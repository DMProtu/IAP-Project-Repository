const PAGE_LIMIT = 500;
const HISTORY_WINDOW_DAYS = 7;

function jsonResponse(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      "Content-Type": "application/json",
      "Cache-Control": "no-store"
    }
  });
}

function getStoredKey(env) {
  return `latest:${env.TTN_DEVICE_ID || "device"}`;
}

function isSensorOnline(receivedAt) {
  if (!receivedAt || receivedAt === "--") return false;

  const lastUpdate = new Date(receivedAt);
  if (Number.isNaN(lastUpdate.getTime())) return false;

  return (new Date() - lastUpdate) / 1000 < 120;
}

function decodeTTNPayload(frmPayload) {
  if (!frmPayload) return null;

  try {
    const binary = atob(frmPayload);
    const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
    const text = new TextDecoder().decode(bytes);

    const parts = text.split(",").map((part) => part.trim());

    const decoded = {
      rawText: text
    };

    for (let i = 0; i < 8; i += 1) {
      const title = parts[i * 2];
      const value = parts[i * 2 + 1];

      decoded[`Title${i + 1}`] = title || `Title${i + 1}`;
      decoded[`Var${i + 1}`] =
        value !== undefined && !Number.isNaN(Number(value))
          ? Number(value)
          : null;
    }

    return decoded;
  } catch {
    return null;
  }
}

function getWebhookRoot(body) {
  if (body?.data?.uplink_message) {
    return body.data;
  }

  return body;
}

function extractWebhookData(body, env) {
  const root = getWebhookRoot(body);
  const endDeviceIds = root?.end_device_ids || {};
  const uplinkMessage = root?.uplink_message || {};

  return {
    device_id: endDeviceIds.device_id || env.TTN_DEVICE_ID || "--",
    received_at:
      root?.received_at ||
      uplinkMessage?.received_at ||
      body?.time ||
      new Date().toISOString(),

    frm_payload: uplinkMessage?.frm_payload || null,
    f_port: uplinkMessage?.f_port ?? "--",
    decoded_payload: uplinkMessage?.decoded_payload || null
  };
}

async function readLatestDeviceData(env) {
  const raw = await env.DEVICE_DATA.get(getStoredKey(env));

  if (!raw) {
    return {
      cloudConnected: true,
      sensorActive: false,
      error: "",
      message: "Connected to webhook backend, but no uplink has been received yet.",
      device_id: env.TTN_DEVICE_ID || "--",
      received_at: "--",
      frm_payload: null,
      f_port: "--",
      decoded_payload: null
    };
  }

  try {
    const parsed = JSON.parse(raw);

    return {
      cloudConnected: true,
      sensorActive: isSensorOnline(parsed.received_at),
      error: "",
      message: isSensorOnline(parsed.received_at)
        ? "Sensor data received."
        : "No recent sensor data received.",
      ...parsed
    };
  } catch {
    return {
      cloudConnected: true,
      sensorActive: false,
      error: "Stored device data is invalid.",
      message: "",
      device_id: env.TTN_DEVICE_ID || "--",
      received_at: "--",
      frm_payload: null,
      f_port: "--",
      decoded_payload: null
    };
  }
}

function parseEventStream(text) {
  const lines = text
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) =>
      line.startsWith("data:")
        ? line.replace(/^data:\s*/, "")
        : line
    );

  const rows = [];

  for (const line of lines) {
    try {
      const parsed = JSON.parse(line);
      rows.push(parsed?.result || parsed);
    } catch {}
  }

  return rows;
}

function normaliseHistoryRow(result) {
  const uplinkMessage = result?.uplink_message || null;

  const receivedAt =
    result?.received_at ||
    uplinkMessage?.received_at ||
    null;

  const frmPayload = uplinkMessage?.frm_payload || null;

  let decoded = null;
  let rawText = "--";

  if (frmPayload) {
    decoded = decodeTTNPayload(frmPayload);

    if (decoded) {
      rawText = decoded.rawText || "--";
    }
  }

  if (!receivedAt) return null;

  return {
    timestamp: receivedAt,

    Title1: decoded?.Title1 || "Title1",
    Var1: decoded?.Var1 ?? null,

    Title2: decoded?.Title2 || "Title2",
    Var2: decoded?.Var2 ?? null,

    Title3: decoded?.Title3 || "Title3",
    Var3: decoded?.Var3 ?? null,

    Title4: decoded?.Title4 || "Title4",
    Var4: decoded?.Var4 ?? null,

    Title5: decoded?.Title5 || "Title5",
    Var5: decoded?.Var5 ?? null,

    Title6: decoded?.Title6 || "Title6",
    Var6: decoded?.Var6 ?? null,

    Title7: decoded?.Title7 || "Title7",
    Var7: decoded?.Var7 ?? null,

    Title8: decoded?.Title8 || "Title8",
    Var8: decoded?.Var8 ?? null,

    rawText
  };
}

async function fetchHistoryPage(env, afterIso, beforeIso) {
  const {
    TTN_REGION: region,
    TTN_APP_ID: appId,
    TTN_API_KEY: apiKey
  } = env;

  const url =
    `https://${region}.cloud.thethings.network/api/v3/as/applications/` +
    `${encodeURIComponent(appId)}/packages/storage/uplink_message` +
    `?after=${encodeURIComponent(afterIso)}` +
    `&before=${encodeURIComponent(beforeIso)}` +
    `&order=-received_at` +
    `&limit=${PAGE_LIMIT}`;

  const response = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      Accept: "text/event-stream"
    }
  });

  if (!response.ok) {
    const text = await response.text();

    throw new Error(
      `TTN history request failed: ${response.status} ${response.statusText}${text ? ` - ${text}` : ""}`
    );
  }

  const body = await response.text();

  return parseEventStream(body);
}

async function fetchHistoryFromTTN(env) {
  const {
    TTN_REGION: region,
    TTN_APP_ID: appId,
    TTN_DEVICE_ID: deviceId,
    TTN_API_KEY: apiKey
  } = env;

  if (!region || !appId || !deviceId || !apiKey) {
    return {
      ok: false,
      status: 500,
      error: "Missing TTN environment variables for history."
    };
  }

  const now = new Date();

  const afterDate = new Date(
    now.getTime() - HISTORY_WINDOW_DAYS * 24 * 60 * 60 * 1000
  );

  let beforeCursor = now.toISOString();

  const history = [];
  const seenKeys = new Set();

  try {
    while (true) {
      const rows = await fetchHistoryPage(
        env,
        afterDate.toISOString(),
        beforeCursor
      );

      if (!rows.length) break;

      let oldestTimestampThisPage = null;
      let deviceRowsThisPage = 0;

      for (const result of rows) {
        const endDeviceIds = result?.end_device_ids || {};

        if (endDeviceIds.device_id !== deviceId) continue;

        const row = normaliseHistoryRow(result);

        if (!row) continue;

        const rowTime = new Date(row.timestamp).getTime();

        if (Number.isNaN(rowTime)) continue;

        deviceRowsThisPage += 1;

        const dedupeKey =
          `${row.timestamp}|${row.rawText}|` +
          `${row.Title1}|${row.Var1}|` +
          `${row.Title2}|${row.Var2}|` +
          `${row.Title3}|${row.Var3}|` +
          `${row.Title4}|${row.Var4}|` +
          `${row.Title5}|${row.Var5}|` +
          `${row.Title6}|${row.Var6}|` +
          `${row.Title7}|${row.Var7}|` +
          `${row.Title8}|${row.Var8}`;

        if (!seenKeys.has(dedupeKey)) {
          seenKeys.add(dedupeKey);
          history.push(row);
        }

        if (
          oldestTimestampThisPage == null ||
          rowTime < new Date(oldestTimestampThisPage).getTime()
        ) {
          oldestTimestampThisPage = row.timestamp;
        }
      }

      if (!oldestTimestampThisPage) {
        break;
      }

      const oldestMs = new Date(oldestTimestampThisPage).getTime();

      if (Number.isNaN(oldestMs)) {
        break;
      }

      if (oldestMs <= afterDate.getTime()) {
        break;
      }

      beforeCursor = new Date(oldestMs - 1).toISOString();

      if (rows.length < PAGE_LIMIT) {
        break;
      }

      if (deviceRowsThisPage === 0) {
        break;
      }
    }

    history.sort(
      (a, b) => new Date(a.timestamp) - new Date(b.timestamp)
    );

    return {
      ok: true,
      status: 200,
      history
    };
  } catch (error) {
    return {
      ok: false,
      status: 500,
      error: `Failed to fetch TTN history: ${error.message}`
    };
  }
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === "/api/health") {
      return jsonResponse({ status: "ok" });
    }

    if (url.pathname === "/api/test-connection") {
      const latest = await readLatestDeviceData(env);

      return jsonResponse({
        status: "ok",
        cloudConnected: true,
        sensorActive: latest.sensorActive,
        message: latest.message
      });
    }

    if (url.pathname === "/api/device-data") {
      return jsonResponse(await readLatestDeviceData(env));
    }

    if (url.pathname === "/api/history") {
      const result = await fetchHistoryFromTTN(env);

      if (!result.ok) {
        return jsonResponse(
          {
            error: result.error,
            history: []
          },
          result.status
        );
      }

      return jsonResponse({
        error: "",
        history: result.history
      });
    }

    if (
      url.pathname === "/api/ttn-webhook" &&
      request.method === "POST"
    ) {
      if (env.WEBHOOK_SECRET) {
        const suppliedSecret =
          request.headers.get("x-webhook-secret") ||
          url.searchParams.get("secret");

        if (suppliedSecret !== env.WEBHOOK_SECRET) {
          return jsonResponse(
            { error: "Unauthorized webhook request." },
            401
          );
        }
      }

      let body;

      try {
        body = await request.json();
      } catch {
        return jsonResponse({ error: "Invalid JSON body." }, 400);
      }

      if (
        !body?.uplink_message &&
        !body?.data?.uplink_message
      ) {
        return jsonResponse({
          status: "ignored",
          message:
            "Request received, but no uplink_message was present."
        });
      }

      await env.DEVICE_DATA.put(
        getStoredKey(env),
        JSON.stringify(extractWebhookData(body, env))
      );

      return jsonResponse({
        status: "stored",
        message: "Uplink stored successfully."
      });
    }

    return env.ASSETS.fetch(request);
  }
};