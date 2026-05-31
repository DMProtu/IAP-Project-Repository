import React, { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

const AEST_OFFSET_MS = 10 * 60 * 60 * 1000;
const WEEK_MS = 7 * 24 * 60 * 60 * 1000;
const DEFAULT_START_PERCENT = (6 / 7) * 100;
const SERIES_GAP_THRESHOLD_MS = 5 * 60 * 1000;

const MONTH_NAMES = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec"
];

const THEME_PRESETS = {
  dark: {
    bgMain: "#2b2f34",
    bgPanel: "#343a40",
    border: "#4b5560",
    textPrimary: "#f5f7fa",
    textSecondary: "#b7c0cb",
    accentBlue: "#27c1e6",
    switchColor: "#27c1e6",
    dangerBg: "#ff63631a",
    dangerBorder: "#ff636359",
    onlineBg: "#33d17a29",
    onlineText: "#86efac",
    offlineBg: "#6b4a4a",
    offlineText: "#ffd6d6",
    softPanel: "#ffffff08",
    softBorder: "#ffffff10",
    sliderRange: "#27c1e673",
    sliderTrack: "#ffffff1f",
    graphGrid: "#ffffff17",
    graphGridVertical: "#ffffff0d",
    graphAxis: "#ffffff38",
    homeHeroStart: "#27c1e629",
    homeHeroEnd: "#343a40f2",
    backButton: "#d84b4b",
    backButtonText: "#fff5f5",
    tooltipBg: "#343a40",
    tooltipBorder: "#4b5560",
    warningBg: "#ffb34724",
    warningBorder: "#ffb3474d",
    warningText: "#ffd59a",
    okBg: "#33d17a24",
    okBorder: "#33d17a47",
    okText: "#9ff0b9",
    scrollbarThumb: "#ffffff1f",
    shadowStrong: "#00000047",
    shadowSoft: "#0000001f",
    sliderThumbRing: "#ffffff1f"
  },

  light: {
    bgMain: "#eef2f7",
    bgPanel: "#ffffff",
    border: "#cfd8e3",
    textPrimary: "#17212b",
    textSecondary: "#5f6f82",
    accentBlue: "#1976d2",
    switchColor: "#1976d2",
    dangerBg: "#dc35451a",
    dangerBorder: "#dc354540",
    onlineBg: "#2ecc7126",
    onlineText: "#1f8f4e",
    offlineBg: "#dc354520",
    offlineText: "#b3261e",
    softPanel: "#17212b08",
    softBorder: "#17212b14",
    sliderRange: "#1976d247",
    sliderTrack: "#17212b1f",
    graphGrid: "#17212b14",
    graphGridVertical: "#17212b0d",
    graphAxis: "#17212b38",
    homeHeroStart: "#1976d21f",
    homeHeroEnd: "#fffffff5",
    backButton: "#cc3d3d",
    backButtonText: "#ffffff",
    tooltipBg: "#ffffff",
    tooltipBorder: "#cfd8e3",
    warningBg: "#ffb34724",
    warningBorder: "#ffb3474d",
    warningText: "#8a5a16",
    okBg: "#2ecc7124",
    okBorder: "#2ecc7147",
    okText: "#1f8f4e",
    scrollbarThumb: "#17212b26",
    shadowStrong: "#17212b29",
    shadowSoft: "#17212b14",
    sliderThumbRing: "#17212b1f"
  },

  redGreenBlind: {
    bgMain: "#1f242b",
    bgPanel: "#28303a",
    border: "#8091a3",
    textPrimary: "#f7fbff",
    textSecondary: "#d3dbe5",

    accentBlue: "#ffd166",
    switchColor: "#ffd166",

    dangerBg: "#7a5cfa24",
    dangerBorder: "#7a5cfa61",

    onlineBg: "#00b4d82e",
    onlineText: "#90e0ef",

    offlineBg: "#7a5cfa2e",
    offlineText: "#e2d7ff",

    softPanel: "#ffffff0d",
    softBorder: "#ffffff1a",

    sliderRange: "#ffd16666",
    sliderTrack: "#ffffff29",

    graphGrid: "#ffffff1f",
    graphGridVertical: "#ffffff12",
    graphAxis: "#ffffff47",

    homeHeroStart: "#ffd1662e",
    homeHeroEnd: "#28303af5",

    backButton: "#ffd166",
    backButtonText: "#1a1a1a",

    tooltipBg: "#28303a",
    tooltipBorder: "#8091a3",

    warningBg: "#ffd16626",
    warningBorder: "#ffd16659",
    warningText: "#ffe29b",

    okBg: "#00b4d826",
    okBorder: "#00b4d85e",
    okText: "#90e0ef",

    scrollbarThumb: "#ffffff26",
    shadowStrong: "#00000047",
    shadowSoft: "#00000024",
    sliderThumbRing: "#ffd16633"
  },

  tritanopia: {
    bgMain: "#20262c",
    bgPanel: "#2a3338",
    border: "#80969a",
    textPrimary: "#f8fcff",
    textSecondary: "#d6dfe2",
    accentBlue: "#00bcd4",
    switchColor: "#00bcd4",
    dangerBg: "#ff4d6d24",
    dangerBorder: "#ff4d6d61",
    onlineBg: "#00bcd42e",
    onlineText: "#b2f5ff",
    offlineBg: "#ff4d6d2e",
    offlineText: "#ffd6de",
    softPanel: "#ffffff0d",
    softBorder: "#ffffff1a",
    sliderRange: "#00bcd466",
    sliderTrack: "#ffffff29",
    graphGrid: "#ffffff1f",
    graphGridVertical: "#ffffff12",
    graphAxis: "#ffffff47",
    homeHeroStart: "#00bcd430",
    homeHeroEnd: "#2a3338f5",
    backButton: "#ff4d6d",
    backButtonText: "#fff5f7",
    tooltipBg: "#2a3338",
    tooltipBorder: "#80969a",
    warningBg: "#ff4d6d26",
    warningBorder: "#ff4d6d59",
    warningText: "#ffc0cb",
    okBg: "#00bcd426",
    okBorder: "#00bcd45e",
    okText: "#b2f5ff",
    scrollbarThumb: "#ffffff26",
    shadowStrong: "#00000047",
    shadowSoft: "#00000024",
    sliderThumbRing: "#ffffff1f"
  }
};

const CUSTOM_THEME_FIELDS = [
  ["bgMain", "Main Background", "color"],
  ["bgPanel", "Panel Background", "color"],
  ["border", "Border", "color"],
  ["textPrimary", "Primary Text", "color"],
  ["textSecondary", "Secondary Text", "color"],
  ["accentBlue", "Accent", "color"],
  ["switchColor", "Switch Active", "color"],
  ["dangerBg", "Danger Background", "color"],
  ["dangerBorder", "Danger Border", "color"],
  ["onlineBg", "Online Background", "color"],
  ["onlineText", "Online Text", "color"],
  ["offlineBg", "Offline Background", "color"],
  ["offlineText", "Offline Text", "color"],
  ["softPanel", "Soft Panel", "color"],
  ["softBorder", "Soft Border", "color"],
  ["sliderRange", "Slider Range", "color"],
  ["sliderTrack", "Slider Track", "color"],
  ["graphGrid", "Graph Grid", "color"],
  ["graphGridVertical", "Graph Grid Vertical", "color"],
  ["graphAxis", "Graph Axis", "color"],
  ["homeHeroStart", "Home Gradient Start", "color"],
  ["homeHeroEnd", "Home Gradient End", "color"],
  ["backButton", "Back Button", "color"],
  ["backButtonText", "Back Button Text", "color"],
  ["tooltipBg", "Tooltip Background", "color"],
  ["tooltipBorder", "Tooltip Border", "color"],
  ["warningBg", "Warning Background", "color"],
  ["warningBorder", "Warning Border", "color"],
  ["warningText", "Warning Text", "color"],
  ["okBg", "Alert Background", "color"],
  ["okBorder", "Alert Border", "color"],
  ["okText", "Alert Text", "color"],
  ["scrollbarThumb", "Scrollbar Thumb", "color"],
  ["shadowStrong", "Strong Shadow", "color"],
  ["shadowSoft", "Soft Shadow", "color"],
  ["sliderThumbRing", "Slider Thumb Ring", "color"]
];

const METRIC_GRAPH_CONFIGS = [
  { key: "Var1", title: "Title1", yTitle: "Degrees Celsius (°C)", exportTitle: "Temperature", yMinDefault: 0, yMaxDefault: 100, yMin: 0, yMax: 100, yStep: 1, yDigits: 0, className: "extra-0", unit: " °C" },
  { key: "Var2", title: "Title2", yTitle: "Turbidity", yMinDefault: 0, yMaxDefault: 100, yMin: 0, yMax: 100, yStep: 1, yDigits: 0, className: "extra-1" },
  { key: "Var3", title: "Title3", yTitle: "Pressure", yMinDefault: 0, yMaxDefault: 100, yMin: 0, yMax: 100, yStep: 1, yDigits: 0, className: "extra-2" },
  { key: "Var4", title: "Title4", yTitle: "Chlorine", yMinDefault: 0, yMaxDefault: 100, yMin: 0, yMax: 100, yStep: 1, yDigits: 0, className: "extra-3" },
  { key: "Var5", title: "Title5", yTitle: "pH", yMinDefault: 0, yMaxDefault: 100, yMin: 0, yMax: 100, yStep: 1, yDigits: 0, className: "extra-4" },
  { key: "Var6", title: "Title6", yTitle: "Conductivity", yMinDefault: 0, yMaxDefault: 100, yMin: 0, yMax: 100, yStep: 1, yDigits: 0, className: "extra-5" },
  { key: "Var7", title: "Title7", yTitle: "Dissolved Oxygen", yMinDefault: 0, yMaxDefault: 100, yMin: 0, yMax: 100, yStep: 1, yDigits: 0, className: "extra-6" },
  { key: "Var8", title: "Title8", yTitle: "Flow", yMinDefault: 0, yMaxDefault: 100, yMin: 0, yMax: 100, yStep: 1, yDigits: 0, className: "extra-7" }
];

const UNIT_OPTIONS = [
  { label: "No unit", value: "" },
  { label: "°C", value: "°C" },
  { label: "°F", value: "°F" },
  { label: "K", value: "K" },
  { label: "V", value: "V" },
  { label: "mV", value: "mV" },
  { label: "A", value: "A" },
  { label: "mA", value: "mA" },
  { label: "Ω", value: "Ω" },
  { label: "kPa", value: "kPa" },
  { label: "Pa", value: "Pa" },
  { label: "pH", value: "pH" },
  { label: "NTU", value: "NTU" },
  { label: "mg/L", value: "mg/L" },
  { label: "L/min", value: "L/min" },
  { label: "%", value: "%" }
];

function App() {
  const [layout, setLayout] = useState("layout2");
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [formatOpen, setFormatOpen] = useState(true);
  const [exportOpen, setExportOpen] = useState(true);
  const [colourOpen, setColourOpen] = useState(true);
  const [page, setPage] = useState("home");
  const [themeMode, setThemeMode] = useState("dark");
  const [customTheme, setCustomTheme] = useState({ ...THEME_PRESETS.dark });

  const [deviceData, setDeviceData] = useState(null);
  const [history, setHistory] = useState([]);
  const [metricUnits, setMetricUnits] = useState({});

  const [rangeStartPercent, setRangeStartPercent] = useState(DEFAULT_START_PERCENT);
  const [rangeEndPercent, setRangeEndPercent] = useState(100);
  const [yRangeMin, setYRangeMin] = useState(20);
  const [yRangeMax, setYRangeMax] = useState(30);

  const [dRangeStartPercent, setDRangeStartPercent] = useState(DEFAULT_START_PERCENT);
  const [dRangeEndPercent, setDRangeEndPercent] = useState(100);
  const [dYRangeMin, setDYRangeMin] = useState(0);
  const [dYRangeMax, setDYRangeMax] = useState(300);

  const [pRangeStartPercent, setPRangeStartPercent] = useState(DEFAULT_START_PERCENT);
  const [pRangeEndPercent, setPRangeEndPercent] = useState(100);
  const [pYRangeMin, setPYRangeMin] = useState(0);
  const [pYRangeMax, setPYRangeMax] = useState(100);

  const [analysisRangeHours, setAnalysisRangeHours] = useState(24);
  const [analysisDayOffset, setAnalysisDayOffset] = useState(0);

  const [alertRangeHours, setAlertRangeHours] = useState(24);
  const [alertThresholds, setAlertThresholds] = useState({
    temperature: 50,
    dValue: 200,
    pressure: 80,
    chlorine: 5,
    ph: 9,
    conductivity: 100,
    dissolvedOxygen: 15,
    flow: 80
  });

  const [liveNow, setLiveNow] = useState(new Date());
  const [cloudStatus, setCloudStatus] = useState("Checking");
  const [sensorStatus, setSensorStatus] = useState("Offline");
  const [errorMessage, setErrorMessage] = useState("");
  const [hoverPoint, setHoverPoint] = useState(null);
  const [waitingDots, setWaitingDots] = useState("");

  const tempGraphAreaRef = useRef(null);
  const dGraphAreaRef = useRef(null);
  const pGraphAreaRef = useRef(null);

  const tempGraphCardRef = useRef(null);
  const dGraphCardRef = useRef(null);
  const pGraphCardRef = useRef(null);

  const [tempGraphSize, setTempGraphSize] = useState({ width: 700, height: 350 });
  const [dGraphSize, setDGraphSize] = useState({ width: 700, height: 350 });
  const [pGraphSize, setPGraphSize] = useState({ width: 700, height: 350 });

  const activeTheme = useMemo(() => {
    if (themeMode === "custom") {
      return { ...THEME_PRESETS.dark, ...customTheme };
    }

    return THEME_PRESETS[themeMode] || THEME_PRESETS.dark;
  }, [themeMode, customTheme]);

  useEffect(() => {
    const root = document.documentElement;

    const vars = {
      "--bg-main": activeTheme.bgMain,
      "--bg-panel": activeTheme.bgPanel,
      "--border": activeTheme.border,
      "--text-primary": activeTheme.textPrimary,
      "--text-secondary": activeTheme.textSecondary,
      "--accent-blue": activeTheme.accentBlue,
      "--switch": activeTheme.switchColor,
      "--danger-bg": activeTheme.dangerBg,
      "--danger-border": activeTheme.dangerBorder,
      "--online-bg": activeTheme.onlineBg,
      "--online-text": activeTheme.onlineText,
      "--offline-bg": activeTheme.offlineBg,
      "--offline-text": activeTheme.offlineText,
      "--soft-panel": activeTheme.softPanel,
      "--soft-border": activeTheme.softBorder,
      "--slider-range": activeTheme.sliderRange,
      "--slider-track": activeTheme.sliderTrack,
      "--graph-grid": activeTheme.graphGrid,
      "--graph-grid-vertical": activeTheme.graphGridVertical,
      "--graph-axis": activeTheme.graphAxis,
      "--home-hero-start": activeTheme.homeHeroStart,
      "--home-hero-end": activeTheme.homeHeroEnd,
      "--back-button": activeTheme.backButton,
      "--back-button-text": activeTheme.backButtonText,
      "--tooltip-bg": activeTheme.tooltipBg,
      "--tooltip-border": activeTheme.tooltipBorder,
      "--warning-bg": activeTheme.warningBg,
      "--warning-border": activeTheme.warningBorder,
      "--warning-text": activeTheme.warningText,
      "--ok-bg": activeTheme.okBg,
      "--ok-border": activeTheme.okBorder,
      "--ok-text": activeTheme.okText,
      "--scrollbar-thumb": activeTheme.scrollbarThumb,
      "--shadow-strong": activeTheme.shadowStrong,
      "--shadow-soft": activeTheme.shadowSoft,
      "--slider-thumb-ring": activeTheme.sliderThumbRing
    };

    Object.entries(vars).forEach(([key, value]) => {
      root.style.setProperty(key, value);
    });
  }, [activeTheme]);

  useEffect(() => {
    const interval = setInterval(() => setLiveNow(new Date()), 1000);
    return () => clearInterval(interval);
  }, []);

  useEffect(() => {
    const frames = ["", ".", "..", "..."];
    let index = 0;

    const interval = setInterval(() => {
      index = (index + 1) % frames.length;
      setWaitingDots(frames[index]);
    }, 500);

    return () => clearInterval(interval);
  }, []);

  useEffect(() => {
    const cleanupTemp = setupGraphResizeObserver(tempGraphAreaRef, setTempGraphSize);
    const cleanupD = setupGraphResizeObserver(dGraphAreaRef, setDGraphSize);
    const cleanupP = setupGraphResizeObserver(pGraphAreaRef, setPGraphSize);

    return () => {
      cleanupTemp();
      cleanupD();
      cleanupP();
    };
  }, []);

  useEffect(() => {
    const loadHistory = async () => {
      try {
        const res = await fetch("/api/history", { cache: "no-store" });
        const data = await res.json();

        if (!res.ok || !Array.isArray(data.history)) return;

        setHistory(
          [...data.history]
            .filter((item) => item?.timestamp)
            .sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp))
        );
      } catch {}
    };

    const fetchData = async () => {
      try {
        const res = await fetch("/api/device-data", { cache: "no-store" });
        const data = await res.json();

        if (!res.ok) {
          setCloudStatus("Offline");
          setSensorStatus("Offline");
          setErrorMessage(data.error || data.message || "Unable to load data.");
          setDeviceData(null);
          return;
        }

        setCloudStatus(data.cloudConnected ? "Online" : "Offline");
        setSensorStatus(data.sensorActive ? "Online" : "Offline");
        setErrorMessage(data.error || "");

        const decoded = decodeTTNPayload(data.frm_payload);
        const nextData = {
          device_id: data.device_id || "--",
          last_update: data.received_at || "--",
          raw_text: decoded?.rawText || "--",

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
          Var8: decoded?.Var8 ?? null
        };

        setDeviceData(nextData);

        if (
          data.sensorActive &&
          decoded &&
          Object.entries(decoded).some(
            ([key, value]) => key !== "rawText" && value != null
          )
        ) {
          setHistory((prev) => {
            if (prev.some((item) => item.timestamp === nextData.last_update)) {
              return prev;
            }

            return [
              ...prev,
              {
                timestamp: nextData.last_update,

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

                rawText: decoded?.rawText ?? "--"
              }
            ].sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
          });
        }
      } catch {
        setCloudStatus("Offline");
        setSensorStatus("Offline");
        setErrorMessage("Network request failed.");
      }
    };

    loadHistory();
    fetchData();

    const interval = setInterval(fetchData, 1000);
    return () => clearInterval(interval);
  }, []);

  const sensorOffline = sensorStatus !== "Online";

  const cleanedHistory = useMemo(() => {
    return removeFirstPointOfEachSeries(history, SERIES_GAP_THRESHOLD_MS);
  }, [history]);

  function renderWaitingText() {
    return (
      <span className="waiting-text">
        <span className="waiting-word">Waiting</span>
        <span className="waiting-dots">{waitingDots}</span>
      </span>
    );
  }

  function formatLiveValue(value, digits = 2, suffix = "") {
    if (sensorOffline || value === "--" || value == null || Number.isNaN(Number(value))) {
      return renderWaitingText();
    }

    return `${Number(value).toFixed(digits)}${suffix}`;
  }

  const liveTemperatureText =
    sensorOffline || deviceData?.temperature === "--"
      ? renderWaitingText()
      : `${Number(deviceData.temperature).toFixed(2)} °C`;

  const liveDValueText =
    sensorOffline || deviceData?.dValue === "--"
      ? renderWaitingText()
      : `${deviceData.dValue}`;

  const livePressureText =
    sensorOffline || deviceData?.pressure === "--"
      ? renderWaitingText()
      : `${Number(deviceData.pressure).toFixed(2)}`;

  const liveChlorineText =
    sensorOffline || deviceData?.chlorine === "--"
      ? renderWaitingText()
      : `${Number(deviceData.chlorine).toFixed(1)}`;

  const livePhText =
    sensorOffline || deviceData?.ph === "--"
      ? renderWaitingText()
      : `${Number(deviceData.ph).toFixed(2)}`;

  const liveConductivityText =
    sensorOffline || deviceData?.conductivity === "--"
      ? renderWaitingText()
      : `${Number(deviceData.conductivity).toFixed(1)}`;

  const liveDissolvedOxygenText =
    sensorOffline || deviceData?.dissolvedOxygen === "--"
      ? renderWaitingText()
      : `${Number(deviceData.dissolvedOxygen).toFixed(2)}`;

  const liveFlowText =
    sensorOffline || deviceData?.flow === "--"
      ? renderWaitingText()
      : `${Number(deviceData.flow).toFixed(1)}`;

  const getLiveMetricText = (config) => {
    const value = deviceData?.[config.key];

    if (sensorOffline || value === "--" || value == null || Number.isNaN(Number(value))) {
      return renderWaitingText();
    }

    return `${Number(value).toFixed(config.yDigits)}${config.unit || ""}`;
  };

  const getMetricTitle = (config) => {
    return deviceData?.[config.title] || config.title;
  };

  const getMetricAxisTitle = (config) => {
    const title = getMetricTitle(config);
    const unit = metricUnits[config.key] || "";

    return unit ? `${title} (${unit})` : title;
  };

  const rawTemperatureText =
    sensorOffline || deviceData?.temperature === "--"
      ? renderWaitingText()
      : `${Number(deviceData.temperature).toFixed(2)} °C`;

  const rawDValueText =
    sensorOffline || deviceData?.dValue === "--"
      ? renderWaitingText()
      : `${deviceData.dValue}`;

  const rawPressureText =
    sensorOffline || deviceData?.pressure === "--"
      ? renderWaitingText()
      : `${Number(deviceData.pressure).toFixed(2)}`;

  const rawChlorineText =
    sensorOffline || deviceData?.chlorine === "--"
      ? renderWaitingText()
      : `${Number(deviceData.chlorine).toFixed(1)}`;

  const rawPhText =
    sensorOffline || deviceData?.ph === "--"
      ? renderWaitingText()
      : `${Number(deviceData.ph).toFixed(2)}`;

  const rawConductivityText =
    sensorOffline || deviceData?.conductivity === "--"
      ? renderWaitingText()
      : `${Number(deviceData.conductivity).toFixed(1)}`;

  const rawDissolvedOxygenText =
    sensorOffline || deviceData?.dissolvedOxygen === "--"
      ? renderWaitingText()
      : `${Number(deviceData.dissolvedOxygen).toFixed(2)}`;

  const rawFlowText =
    sensorOffline || deviceData?.flow === "--"
      ? renderWaitingText()
      : `${Number(deviceData.flow).toFixed(1)}`;

  const domainEnd = useMemo(() => new Date(liveNow), [liveNow]);
  const domainStart = useMemo(() => new Date(domainEnd.getTime() - WEEK_MS), [domainEnd]);

  const selectedRange = useMemo(() => {
    const totalMs = domainEnd.getTime() - domainStart.getTime();
    const startMs = domainStart.getTime() + (rangeStartPercent / 100) * totalMs;
    const endMs = domainStart.getTime() + (rangeEndPercent / 100) * totalMs;

    return {
      start: new Date(startMs),
      end: new Date(endMs)
    };
  }, [domainStart, domainEnd, rangeStartPercent, rangeEndPercent]);

  const selectedDRange = useMemo(() => {
    const totalMs = domainEnd.getTime() - domainStart.getTime();
    const startMs = domainStart.getTime() + (dRangeStartPercent / 100) * totalMs;
    const endMs = domainStart.getTime() + (dRangeEndPercent / 100) * totalMs;

    return {
      start: new Date(startMs),
      end: new Date(endMs)
    };
  }, [domainStart, domainEnd, dRangeStartPercent, dRangeEndPercent]);

  const selectedPRange = useMemo(() => {
    const totalMs = domainEnd.getTime() - domainStart.getTime();
    const startMs = domainStart.getTime() + (pRangeStartPercent / 100) * totalMs;
    const endMs = domainStart.getTime() + (pRangeEndPercent / 100) * totalMs;

    return {
      start: new Date(startMs),
      end: new Date(endMs)
    };
  }, [domainStart, domainEnd, pRangeStartPercent, pRangeEndPercent]);

  const visibleTempHistory = useMemo(() => {
    const startMs = selectedRange.start.getTime();
    const endMs = selectedRange.end.getTime();

    return cleanedHistory.filter((item) => {
      const t = new Date(item.timestamp).getTime();
      const temp = Number(item.temperature);

      return (
        !Number.isNaN(t) &&
        !Number.isNaN(temp) &&
        t >= startMs &&
        t <= endMs &&
        temp >= yRangeMin &&
        temp <= yRangeMax
      );
    });
  }, [cleanedHistory, selectedRange, yRangeMin, yRangeMax]);

  const visibleDHistory = useMemo(() => {
    const startMs = selectedDRange.start.getTime();
    const endMs = selectedDRange.end.getTime();

    return cleanedHistory.filter((item) => {
      const t = new Date(item.timestamp).getTime();
      const dValue = Number(item.dValue);

      return (
        !Number.isNaN(t) &&
        !Number.isNaN(dValue) &&
        t >= startMs &&
        t <= endMs &&
        dValue >= dYRangeMin &&
        dValue <= dYRangeMax
      );
    });
  }, [cleanedHistory, selectedDRange, dYRangeMin, dYRangeMax]);

  const visiblePHistory = useMemo(() => {
    const startMs = selectedPRange.start.getTime();
    const endMs = selectedPRange.end.getTime();

    return cleanedHistory.filter((item) => {
      const t = new Date(item.timestamp).getTime();
      const pressure = Number(item.pressure);

      return (
        !Number.isNaN(t) &&
        !Number.isNaN(pressure) &&
        t >= startMs &&
        t <= endMs &&
        pressure >= pYRangeMin &&
        pressure <= pYRangeMax
      );
    });
  }, [cleanedHistory, selectedPRange, pYRangeMin, pYRangeMax]);

  const analysisHistory = useMemo(() => {
    const endMs = domainEnd.getTime();
    const startMs = endMs - analysisRangeHours * 60 * 60 * 1000;

    return cleanedHistory.filter((item) => {
      const t = new Date(item.timestamp).getTime();
      return !Number.isNaN(t) && t >= startMs && t <= endMs;
    });
  }, [cleanedHistory, domainEnd, analysisRangeHours]);

  const alertHistory = useMemo(() => {
    const endMs = domainEnd.getTime();
    const startMs = endMs - alertRangeHours * 60 * 60 * 1000;

    return cleanedHistory.filter((item) => {
      const t = new Date(item.timestamp).getTime();
      return !Number.isNaN(t) && t >= startMs && t <= endMs;
    });
  }, [cleanedHistory, domainEnd, alertRangeHours]);

  const metricStats = useMemo(() => {
    return Object.fromEntries(
      METRIC_GRAPH_CONFIGS.map((config) => [
        config.key,
        getSeriesStats(analysisHistory, config.key)
      ])
    );
  }, [analysisHistory]);

  const analysisDayOptions = useMemo(
    () => [
      { value: 0, label: "Today" },
      { value: 1, label: "1 day ago" },
      { value: 2, label: "2 days ago" },
      { value: 3, label: "3 days ago" },
      { value: 4, label: "4 days ago" },
      { value: 5, label: "5 days ago" },
      { value: 6, label: "6 days ago" }
    ],
    []
  );

  const nightSeries = useMemo(
    () => filterSeriesByTimeWindow(cleanedHistory, analysisDayOffset, 0, 6),
    [cleanedHistory, analysisDayOffset]
  );
  const morningSeries = useMemo(
    () => filterSeriesByTimeWindow(cleanedHistory, analysisDayOffset, 6, 12),
    [cleanedHistory, analysisDayOffset]
  );
  const afternoonSeries = useMemo(
    () => filterSeriesByTimeWindow(cleanedHistory, analysisDayOffset, 12, 18),
    [cleanedHistory, analysisDayOffset]
  );
  const eveningSeries = useMemo(
    () => filterSeriesByTimeWindow(cleanedHistory, analysisDayOffset, 18, 24),
    [cleanedHistory, analysisDayOffset]
  );

  const nightMetricStats = useMemo(
    () => getMetricStatsForSeries(nightSeries),
    [nightSeries]
  );

  const morningMetricStats = useMemo(
    () => getMetricStatsForSeries(morningSeries),
    [morningSeries]
  );

  const afternoonMetricStats = useMemo(
    () => getMetricStatsForSeries(afternoonSeries),
    [afternoonSeries]
  );

  const eveningMetricStats = useMemo(
    () => getMetricStatsForSeries(eveningSeries),
    [eveningSeries]
  );

  const alertEvents = useMemo(
    () =>
      buildAlertEvents({
        series: alertHistory,
        thresholds: alertThresholds,
        configs: METRIC_GRAPH_CONFIGS.map((config) => ({
          ...config,
          alertTitle: getMetricAxisTitle(config),
          selectedUnit: metricUnits[config.key] || ""
        }))
      }),
    [alertHistory, alertThresholds, metricUnits, deviceData]
  );

  const metricAlerts = useMemo(() => {
    return Object.fromEntries(
      METRIC_GRAPH_CONFIGS.map((config) => [
        config.key,
        alertEvents.filter((alert) => alert.type === config.key)
      ])
    );
  }, [alertEvents]);

  const rangeLabels = useMemo(
    () => ({
      left: formatGraphTimeAEST(domainStart),
      right: formatGraphTimeAEST(domainEnd)
    }),
    [domainStart, domainEnd]
  );

  const dRangeLabels = useMemo(
    () => ({
      left: formatGraphTimeAEST(domainStart),
      right: formatGraphTimeAEST(domainEnd)
    }),
    [domainStart, domainEnd]
  );

  const pRangeLabels = useMemo(
    () => ({
      left: formatGraphTimeAEST(domainStart),
      right: formatGraphTimeAEST(domainEnd)
    }),
    [domainStart, domainEnd]
  );

  const manualStartParts = useMemo(() => toAestParts(selectedRange.start), [selectedRange.start]);
  const manualEndParts = useMemo(() => toAestParts(selectedRange.end), [selectedRange.end]);
  const manualDStartParts = useMemo(() => toAestParts(selectedDRange.start), [selectedDRange.start]);
  const manualDEndParts = useMemo(() => toAestParts(selectedDRange.end), [selectedDRange.end]);
  const manualPStartParts = useMemo(() => toAestParts(selectedPRange.start), [selectedPRange.start]);
  const manualPEndParts = useMemo(() => toAestParts(selectedPRange.end), [selectedPRange.end]);

  const yearOptions = useMemo(() => {
    const startYear = toAestParts(domainStart).year;
    const endYear = toAestParts(domainEnd).year;
    const years = [];

    for (let year = startYear; year <= endYear; year += 1) {
      years.push(year);
    }

    return years;
  }, [domainStart, domainEnd]);

  const yOptions = useMemo(() => Array.from({ length: 101 }, (_, i) => i), []);
  const dYOptions = useMemo(() => Array.from({ length: 301 }, (_, i) => i), []);
  const pYOptions = useMemo(() => Array.from({ length: 101 }, (_, i) => i), []);

  const handleManualTimeChange = (which, field, rawValue) => {
    const parts = which === "start" ? { ...manualStartParts } : { ...manualEndParts };

    parts[field] = ["day", "month", "year", "hour", "minute", "second"].includes(field)
      ? Number(rawValue)
      : rawValue;

    let nextMs = aestPartsToDate(parts).getTime();
    const minMs = domainStart.getTime();
    const maxMs = domainEnd.getTime();
    const otherMs =
      which === "start" ? selectedRange.end.getTime() : selectedRange.start.getTime();

    nextMs = clamp(nextMs, minMs, maxMs);
    if (which === "start") nextMs = Math.min(nextMs, otherMs);
    if (which === "end") nextMs = Math.max(nextMs, otherMs);

    const totalMs = Math.max(maxMs - minMs, 1);
    const nextPercent = ((nextMs - minMs) / totalMs) * 100;

    if (which === "start") {
      setRangeStartPercent(Math.min(nextPercent, rangeEndPercent));
    } else {
      setRangeEndPercent(Math.max(nextPercent, rangeStartPercent));
    }
  };

  const handleDManualTimeChange = (which, field, rawValue) => {
    const parts = which === "start" ? { ...manualDStartParts } : { ...manualDEndParts };

    parts[field] = ["day", "month", "year", "hour", "minute", "second"].includes(field)
      ? Number(rawValue)
      : rawValue;

    let nextMs = aestPartsToDate(parts).getTime();
    const minMs = domainStart.getTime();
    const maxMs = domainEnd.getTime();
    const otherMs =
      which === "start" ? selectedDRange.end.getTime() : selectedDRange.start.getTime();

    nextMs = clamp(nextMs, minMs, maxMs);
    if (which === "start") nextMs = Math.min(nextMs, otherMs);
    if (which === "end") nextMs = Math.max(nextMs, otherMs);

    const totalMs = Math.max(maxMs - minMs, 1);
    const nextPercent = ((nextMs - minMs) / totalMs) * 100;

    if (which === "start") {
      setDRangeStartPercent(Math.min(nextPercent, dRangeEndPercent));
    } else {
      setDRangeEndPercent(Math.max(nextPercent, dRangeStartPercent));
    }
  };

  const handlePManualTimeChange = (which, field, rawValue) => {
    const parts = which === "start" ? { ...manualPStartParts } : { ...manualPEndParts };

    parts[field] = ["day", "month", "year", "hour", "minute", "second"].includes(field)
      ? Number(rawValue)
      : rawValue;

    let nextMs = aestPartsToDate(parts).getTime();
    const minMs = domainStart.getTime();
    const maxMs = domainEnd.getTime();
    const otherMs =
      which === "start" ? selectedPRange.end.getTime() : selectedPRange.start.getTime();

    nextMs = clamp(nextMs, minMs, maxMs);
    if (which === "start") nextMs = Math.min(nextMs, otherMs);
    if (which === "end") nextMs = Math.max(nextMs, otherMs);

    const totalMs = Math.max(maxMs - minMs, 1);
    const nextPercent = ((nextMs - minMs) / totalMs) * 100;

    if (which === "start") {
      setPRangeStartPercent(Math.min(nextPercent, pRangeEndPercent));
    } else {
      setPRangeEndPercent(Math.max(nextPercent, pRangeStartPercent));
    }
  };

  const handleYManualChange = (which, rawValue) => {
    const value = clamp(Number(rawValue), 0, 100);

    if (which === "min") {
      setYRangeMin(Math.min(value, yRangeMax));
    } else {
      setYRangeMax(Math.max(value, yRangeMin));
    }
  };

  const handleDYManualChange = (which, rawValue) => {
    const value = clamp(Number(rawValue), 0, 300);

    if (which === "min") {
      setDYRangeMin(Math.min(value, dYRangeMax));
    } else {
      setDYRangeMax(Math.max(value, dYRangeMin));
    }
  };

  const handlePYManualChange = (which, rawValue) => {
    const value = clamp(Number(rawValue), 0, 100);

    if (which === "min") {
      setPYRangeMin(Math.min(value, pYRangeMax));
    } else {
      setPYRangeMax(Math.max(value, pYRangeMin));
    }
  };

  const tempGraphData = useMemo(
    () =>
      getGraphData({
        series: visibleTempHistory,
        xStart: selectedRange.start,
        xEnd: selectedRange.end,
        yMin: yRangeMin,
        yMax: yRangeMax,
        width: tempGraphSize.width,
        height: tempGraphSize.height,
        valueKey: "temperature",
        xTitle: "Time",
        yTitle: "Temperature"
      }),
    [visibleTempHistory, selectedRange, yRangeMin, yRangeMax, tempGraphSize]
  );

  const dGraphData = useMemo(
    () =>
      getGraphData({
        series: visibleDHistory,
        xStart: selectedDRange.start,
        xEnd: selectedDRange.end,
        yMin: dYRangeMin,
        yMax: dYRangeMax,
        width: dGraphSize.width,
        height: dGraphSize.height,
        valueKey: "dValue",
        xTitle: "Time",
        yTitle: "Turbidity"
      }),
    [visibleDHistory, selectedDRange, dYRangeMin, dYRangeMax, dGraphSize]
  );

  const pGraphData = useMemo(
    () =>
      getGraphData({
        series: visiblePHistory,
        xStart: selectedPRange.start,
        xEnd: selectedPRange.end,
        yMin: pYRangeMin,
        yMax: pYRangeMax,
        width: pGraphSize.width,
        height: pGraphSize.height,
        valueKey: "pressure",
        xTitle: "Time",
        yTitle: "Pressure"
      }),
    [visiblePHistory, selectedPRange, pYRangeMin, pYRangeMax, pGraphSize]
  );

  const exportExcelFile = (rows, fileName) => {
    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Data");
    XLSX.writeFile(workbook, fileName);
  };

  const exportData = () => {
    const rows = cleanedHistory.map((item) => ({
      timestamp: formatTimestamp(item.timestamp),
      iso_timestamp: item.timestamp,

      [deviceData?.Title1 || "Title1"]: item.Var1 ?? "",
      [deviceData?.Title2 || "Title2"]: item.Var2 ?? "",
      [deviceData?.Title3 || "Title3"]: item.Var3 ?? "",
      [deviceData?.Title4 || "Title4"]: item.Var4 ?? "",
      [deviceData?.Title5 || "Title5"]: item.Var5 ?? "",
      [deviceData?.Title6 || "Title6"]: item.Var6 ?? "",
      [deviceData?.Title7 || "Title7"]: item.Var7 ?? "",
      [deviceData?.Title8 || "Title8"]: item.Var8 ?? ""
    }));

    exportExcelFile(rows, "graph-data.xlsx");
  };

  const exportGraphPng = async (ref, fileName) => {
    const graphArea = ref?.current;
    if (!graphArea) return;

    const svg = graphArea.querySelector("svg");
    if (!svg) return;

    try {
      const graphRect = graphArea.getBoundingClientRect();
      const svgRect = svg.getBoundingClientRect();
      const graphStyles = window.getComputedStyle(graphArea);
      const scale = 3;

      const canvas = document.createElement("canvas");
      canvas.width = Math.max(Math.round(graphRect.width * scale), 1);
      canvas.height = Math.max(Math.round(graphRect.height * scale), 1);

      const ctx = canvas.getContext("2d");
      ctx.scale(scale, scale);
      ctx.fillStyle = graphStyles.backgroundColor || activeTheme.bgPanel;
      ctx.fillRect(0, 0, graphRect.width, graphRect.height);

      const clonedSvg = svg.cloneNode(true);
      clonedSvg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
      clonedSvg.setAttribute("width", String(svgRect.width));
      clonedSvg.setAttribute("height", String(svgRect.height));

      inlineSvgStyles(svg, clonedSvg);

      const svgText = new XMLSerializer().serializeToString(clonedSvg);
      const svgBlob = new Blob([svgText], {
        type: "image/svg+xml;charset=utf-8"
      });
      const svgUrl = URL.createObjectURL(svgBlob);

      const image = new Image();
      image.onload = () => {
        const offsetX = svgRect.left - graphRect.left;
        const offsetY = svgRect.top - graphRect.top;

        ctx.drawImage(image, offsetX, offsetY, svgRect.width, svgRect.height);
        URL.revokeObjectURL(svgUrl);

        const link = document.createElement("a");
        link.download = fileName;
        link.href = canvas.toDataURL("image/png");
        link.click();
      };

      image.onerror = () => {
        URL.revokeObjectURL(svgUrl);
        console.error("PNG export failed: SVG could not be rendered.");
      };

      image.src = svgUrl;
    } catch (error) {
      console.error("PNG export failed:", error);
    }
  };

  const metricGraphRefs = useRef({});

  const exportMetricPng = (config) => {
    const ref = metricGraphRefs.current[config.key];

    if (!ref) return;

    exportGraphPng(
      { current: ref },
      `${deviceData?.[config.title] || config.title}-graph.png`
    );
  };

  if (page === "home") {
    return (
      <div className="home-screen">
        <div className="home-card">
          <h1 className="home-title">Remote Water Condition Monitoring Dashboard</h1>
          <p className="home-description">
            This webpage monitors the temperature, turbidity, and water pressure of a body of water in the Gold Coast. This Dashboard displays the live data from remote sensors through a dashboard with live graphs, analysis panels, alerts, and export tools.
          </p>
          <button
            type="button"
            className="home-connect-button"
            onClick={() => setPage("dashboard")}
          >
            Connect
          </button>
          <p className="home-description">
            Made by Daniel McAtamney
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="app">
      <div className={`dashboard-layout ${sidebarOpen ? "sidebar-open" : "sidebar-closed"}`}>
        <button
          type="button"
          className={`sidebar-tab ${sidebarOpen ? "open" : "closed"}`}
          onClick={() => setSidebarOpen((prev) => !prev)}
        >
          ⚙
        </button>

        <aside className={`sidebar ${sidebarOpen ? "open" : "closed"}`}>
          <div className="sidebar-inner">
            <div className="sidebar-title">Settings</div>

            <div className={`sidebar-expandable ${formatOpen ? "open" : ""}`}>
              <button
                type="button"
                className="sidebar-item sidebar-button"
                onClick={() => setFormatOpen((prev) => !prev)}
              >
                <span>Format</span>
                <span className={`chevron ${formatOpen ? "open" : ""}`}>▾</span>
              </button>

              {formatOpen && (
                <div className="format-panel">
                  <FormatToggle
                    label="List"
                    active={layout === "layout1"}
                    onClick={() => setLayout("layout1")}
                  />
                  <FormatToggle
                    label="2 x 3 Grid"
                    active={layout === "layout2"}
                    onClick={() => setLayout("layout2")}
                  />
                  <FormatToggle
                    label="3 x 3 Grid"
                    active={layout === "layout3"}
                    onClick={() => setLayout("layout3")}
                  />
                </div>
              )}
            </div>

            <div className={`sidebar-expandable ${colourOpen ? "open" : ""}`}>
              <button
                type="button"
                className="sidebar-item sidebar-button"
                onClick={() => setColourOpen((prev) => !prev)}
              >
                <span>Webpage Colours</span>
                <span className={`chevron ${colourOpen ? "open" : ""}`}>▾</span>
              </button>

              {colourOpen && (
                <div className="format-panel">
                  <FormatToggle
                    label="Dark Mode"
                    active={themeMode === "dark"}
                    onClick={() => setThemeMode("dark")}
                  />
                  <FormatToggle
                    label="Light Mode"
                    active={themeMode === "light"}
                    onClick={() => setThemeMode("light")}
                  />
                  <FormatToggle
                    label="Protanopia & Deuteranopia"
                    active={themeMode === "redGreenBlind"}
                    onClick={() => setThemeMode("redGreenBlind")}
                  />
                  <FormatToggle
                    label="Tritanopia"
                    active={themeMode === "tritanopia"}
                    onClick={() => setThemeMode("tritanopia")}
                  />
                  <FormatToggle
                    label="Custom"
                    active={themeMode === "custom"}
                    onClick={() => setThemeMode("custom")}
                  />

                  {themeMode === "custom" && (
                    <div className="custom-theme-panel">
                      {CUSTOM_THEME_FIELDS.map(([key, label]) => (
                        <label key={key} className="custom-theme-row">
                          <span>{label}</span>
                          <input
                            type="color"
                            value={isHexColor(customTheme[key]) ? customTheme[key] : "#000000"}
                            onChange={(e) => {
                              setThemeMode("custom");
                              setCustomTheme((prev) => ({
                                ...prev,
                                [key]: e.target.value
                              }));
                            }}
                          />
                        </label>
                      ))}
                    </div>
                  )}
                </div>
              )}
            </div>

            <div className={`sidebar-expandable ${exportOpen ? "open" : ""}`}>
              <button
                type="button"
                className="sidebar-item sidebar-button"
                onClick={() => setExportOpen((prev) => !prev)}
              >
                <span>Export</span>
                <span className={`chevron ${exportOpen ? "open" : ""}`}>▾</span>
              </button>

              {exportOpen && (
                <div className="format-panel">
                  <button
                    type="button"
                    className="sidebar-item sidebar-button"
                    onClick={exportData}
                  >
                    <span>Export Data Excel</span>
                  </button>

                  {METRIC_GRAPH_CONFIGS.map((config) => (
                    <button
                      key={`export-${config.key}`}
                      type="button"
                      className="sidebar-item sidebar-button"
                      onClick={() => exportMetricPng(config)}
                    >
                      <span>{deviceData?.[config.title] || config.title} PNG</span>
                    </button>
                  ))}
                </div>
              )}
            </div>
          </div>

          <div className="sidebar-footer">
            <button
              type="button"
              className="back-button"
              onClick={() => setPage("home")}
            >
              Back
            </button>
          </div>
        </aside>

        <main className="dashboard-main">
          <div className="dashboard-topbar">
            <h1 className="dashboard-heading">Dashboard</h1>

            <div className="connection-panel-card">
              <div className="connection-title">Connection</div>

              <div className="connection-row">
                <span>Cloud</span>
                <span className={`status-pill ${cloudStatus === "Online" ? "online" : "offline"}`}>
                  {cloudStatus}
                </span>
              </div>

              <div className="connection-row">
                <span>Sensor</span>
                <span className={`status-pill ${sensorStatus === "Online" ? "online" : "offline"}`}>
                  {sensorStatus}
                </span>
              </div>
            </div>
          </div>

          {errorMessage ? <div className="error-banner">{errorMessage}</div> : null}

          <div className={`dashboard-content ${layout}`}>
            {METRIC_GRAPH_CONFIGS.map((config) => (
              <MetricGraphPanel
                key={config.key}
                config={config}
                titleText={getMetricTitle(config)}
                axisTitle={getMetricAxisTitle(config)}
                selectedUnit={metricUnits[config.key] || ""}
                onUnitChange={(unit) =>  setMetricUnits((prev) => ({  ...prev,[config.key]: unit}))}
                liveText={getLiveMetricText(config)}
                cleanedHistory={cleanedHistory}
                domainStart={domainStart}
                domainEnd={domainEnd}
                yearOptions={yearOptions}
                hoverPoint={hoverPoint}
                setHoverPoint={setHoverPoint}
                cardRef={(el) => {
                  metricGraphRefs.current[config.key] = el;
                }}
              />
            ))}

            <section className="data-card analysis-panel">
              <div className="card-header">
                <h2>Analysis</h2>

                <div className="analysis-range-selector">
                  <FormatToggle
                    label="1 Hour"
                    active={analysisRangeHours === 1}
                    onClick={() => setAnalysisRangeHours(1)}
                  />
                  <FormatToggle
                    label="24 Hours"
                    active={analysisRangeHours === 24}
                    onClick={() => setAnalysisRangeHours(24)}
                  />
                  <FormatToggle
                    label="7 Days"
                    active={analysisRangeHours === 168}
                    onClick={() => setAnalysisRangeHours(168)}
                  />
                </div>
              </div>

              <div className="analysis-layout analysis-layout-single">
                <div className="analysis-left">
                  {METRIC_GRAPH_CONFIGS.map((config) => (
                    <SummaryBlock
                      key={`summary-${config.key}`}
                      title={`${getMetricAxisTitle(config)} Summary`}
                      stats={metricStats[config.key]}
                      digits={config.yDigits}
                      suffix={metricUnits[config.key] ? ` ${metricUnits[config.key]}` : ""}
                    />
                  ))}
                </div>
              </div>

              <div className="analysis-breakdown-header">
                <h3 className="analysis-subheading">Full Day Analysis Breakdown</h3>

                <div className="analysis-day-selector">
                  <span className="analysis-day-label">Day:</span>
                  <select
                    className="graph-select"
                    value={analysisDayOffset}
                    onChange={(e) => setAnalysisDayOffset(Number(e.target.value))}
                  >
                    {analysisDayOptions.map((option) => (
                      <option key={option.value} value={option.value}>
                        {option.label}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="period-analysis-grid">
                <PeriodAnalysisBlock
                  title="12am – 6am"
                  metricStats={nightMetricStats}
                  getMetricAxisTitle={getMetricAxisTitle}
                  metricUnits={metricUnits}
                />
                <PeriodAnalysisBlock
                  title="6am – 12pm"
                  metricStats={morningMetricStats}
                  getMetricAxisTitle={getMetricAxisTitle}
                  metricUnits={metricUnits}
                />
                <PeriodAnalysisBlock
                  title="12pm – 6pm"
                  metricStats={afternoonMetricStats}
                  getMetricAxisTitle={getMetricAxisTitle}
                  metricUnits={metricUnits}
                />
                <PeriodAnalysisBlock
                  title="6pm – 12am"
                  metricStats={eveningMetricStats}
                  getMetricAxisTitle={getMetricAxisTitle}
                  metricUnits={metricUnits}
                />
              </div>
            </section>

            <section className="data-card alerts-panel">
              <div className="card-header">
                <h2>Alerts</h2>

                <div className="analysis-range-selector">
                  <FormatToggle
                    label="1 Hour"
                    active={alertRangeHours === 1}
                    onClick={() => setAlertRangeHours(1)}
                  />
                  <FormatToggle
                    label="24 Hours"
                    active={alertRangeHours === 24}
                    onClick={() => setAlertRangeHours(24)}
                  />
                  <FormatToggle
                    label="7 Days"
                    active={alertRangeHours === 168}
                    onClick={() => setAlertRangeHours(168)}
                  />
                </div>
              </div>

              <div className="alerts-layout alerts-layout-top">
                {METRIC_GRAPH_CONFIGS.map((config) => {
                  const alerts = metricAlerts[config.key] || [];

                  return (
                    <div key={`alert-box-${config.key}`} className="alert-event-box">
                      <div className="analysis-summary-block">
                        <h3>{getMetricAxisTitle(config)} Events</h3>
                        <div className="analysis-alerts">
                          {alerts.length ? (
                            alerts.map((alert, index) => (
                              <div
                                key={`${config.key}-${index}-${alert.message}`}
                                className="analysis-alert warning"
                              >
                                {alert.message}
                              </div>
                            ))
                          ) : (
                            <div className="analysis-alert ok">
                              No {getMetricAxisTitle(config).toLowerCase()} alerts
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>

              <div className="alerts-layout alerts-layout-bottom">
                <div className="analysis-summary-block alerts-ok-banner">
                  <h3>System Status</h3>
                  <div className="analysis-alerts alerts-status-row">
                    {alertEvents.length === 0 ? (
                      <div className="analysis-alert ok">
                        No alerts triggered in the selected period.
                      </div>
                    ) : null}
                  </div>
                </div>

                <div className="analysis-summary-block threshold-panel">
                  <h3>Alert Thresholds</h3>

                  {METRIC_GRAPH_CONFIGS.map((config) => (
                    <ThresholdSlider
                      key={`threshold-${config.key}`}
                      label={`${getMetricAxisTitle(config)} Threshold`}
                      min={config.yMin}
                      max={config.yMax}
                      step={config.yStep}
                      value={alertThresholds[config.key]}
                      onChange={(value) =>
                        setAlertThresholds((prev) => ({
                          ...prev,
                          [config.key]: value
                        }))
                      }
                      displayValue={`${Number(alertThresholds[config.key]).toFixed(config.yDigits)}` + `${metricUnits[config.key] ? ` ${metricUnits[config.key]}` : ""}`}
                    />
                  ))}
                </div>
              </div>
            </section>

            <section className="data-card raw-data-card">
              <div className="card-header">
                <h2>Raw Data</h2>
                <span className="mini-stat">
                  Last update: {formatTimestamp(deviceData?.last_update)}
                </span>
              </div>

              <div className="data-grid">
                <div className="data-row">
                  <span className="label">Device ID</span>
                  <span className="value">{deviceData?.device_id || "--"}</span>
                </div>

                {Array.from({ length: 8 }, (_, index) => {
                  const titleKey = `Title${index + 1}`;
                  const valueKey = `Var${index + 1}`;

                  return (
                    <div className="data-row" key={valueKey}>
                      <span className="label">
                        {deviceData?.[titleKey] || titleKey}
                      </span>

                      <span className="value">
                        {sensorOffline ||
                        deviceData?.[valueKey] == null
                          ? renderWaitingText()
                          : deviceData[valueKey]}
                      </span>
                    </div>
                  );
                })}

                <div className="data-row">
                  <span className="label">Payload</span>
                  <span className="value raw-text">{deviceData?.raw_text || "--"}</span>
                </div>
              </div>
            </section>
          </div>
        </main>
      </div>
    </div>
  );
}

function FormatToggle({ label, active, onClick }) {
  return (
    <div className="format-option">
      <div className="format-text">
        <span className="format-label">{label}</span>
      </div>

      <button type="button" className={`switch ${active ? "active" : ""}`} onClick={onClick}>
        <span className="switch-knob" />
      </button>
    </div>
  );
}

function MetricGraphPanel({
  config,
  titleText,
  axisTitle,
  selectedUnit,
  onUnitChange,
  liveText,
  cleanedHistory,
  domainStart,
  domainEnd,
  yearOptions,
  hoverPoint,
  setHoverPoint,
  cardRef
}) {

  const [rangeStartPercent, setRangeStartPercent] = useState(DEFAULT_START_PERCENT);
  const [rangeEndPercent, setRangeEndPercent] = useState(100);
  const [yRangeMin, setYRangeMin] = useState(config.yMinDefault);
  const [yRangeMax, setYRangeMax] = useState(config.yMaxDefault);
  const [graphSize, setGraphSize] = useState({ width: 700, height: 350 });
  const [pausedSnapshot, setPausedSnapshot] = useState(null);
  const graphAreaRef = useRef(null);

  useEffect(() => setupGraphResizeObserver(graphAreaRef, setGraphSize), []);

  useEffect(() => {
    if (cardRef && graphAreaRef.current) {
      cardRef(graphAreaRef.current);
    }

    return () => {
      if (cardRef) cardRef(null);
    };
  }, [cardRef]);

  const selectedRange = useMemo(() => {
    const totalMs = domainEnd.getTime() - domainStart.getTime();
    const startMs = domainStart.getTime() + (rangeStartPercent / 100) * totalMs;
    const endMs = domainStart.getTime() + (rangeEndPercent / 100) * totalMs;
    return { start: new Date(startMs), end: new Date(endMs) };
  }, [domainStart, domainEnd, rangeStartPercent, rangeEndPercent]);

  const isPaused = Boolean(pausedSnapshot);

  const activeHistory = isPaused ? pausedSnapshot.history : cleanedHistory;
  const activeSelectedRange = isPaused ? pausedSnapshot.selectedRange : selectedRange;
  const activeYRangeMin = isPaused ? pausedSnapshot.yRangeMin : yRangeMin;
  const activeYRangeMax = isPaused ? pausedSnapshot.yRangeMax : yRangeMax;
  const activeLiveText = isPaused ? pausedSnapshot.liveText : liveText;

  const visibleHistory = useMemo(() => {
    const startMs = activeSelectedRange.start.getTime();
    const endMs = activeSelectedRange.end.getTime();

    return activeHistory.filter((item) => {
      const t = new Date(item.timestamp).getTime();
      const value = Number(item[config.key]);

      return (
        !Number.isNaN(t) &&
        !Number.isNaN(value) &&
        t >= startMs &&
        t <= endMs &&
        value >= activeYRangeMin &&
        value <= activeYRangeMax
      );
    });
  }, [activeHistory, activeSelectedRange, activeYRangeMin, activeYRangeMax, config.key]);

  const graphData = useMemo(
    () => getGraphData({
      series: visibleHistory,
      xStart: activeSelectedRange.start,
      xEnd: activeSelectedRange.end,
      yMin: activeYRangeMin,
      yMax: activeYRangeMax,
      width: graphSize.width,
      height: graphSize.height,
      valueKey: config.key,
      xTitle: "Time",
      yTitle: axisTitle
    }),
    [
      visibleHistory,
      activeSelectedRange,
      activeYRangeMin,
      activeYRangeMax,
      graphSize,
      config.key,
      config.yTitle
    ]
  );

  const rangeLabels = useMemo(() => ({ left: formatGraphTimeAEST(domainStart), right: formatGraphTimeAEST(domainEnd) }), [domainStart, domainEnd]);
  const manualStartParts = useMemo(() => toAestParts(selectedRange.start), [selectedRange.start]);
  const manualEndParts = useMemo(() => toAestParts(selectedRange.end), [selectedRange.end]);

  const yOptions = useMemo(() => {
    const values = [];
    const decimals = String(config.yStep).includes(".") ? String(config.yStep).split(".")[1].length : 0;
    for (let value = config.yMin; value <= config.yMax + config.yStep / 2; value += config.yStep) {
      values.push(Number(value.toFixed(decimals)));
    }
    return values;
  }, [config.yMin, config.yMax, config.yStep]);

  const handlePause = () => {
    setPausedSnapshot({
      history: cleanedHistory,
      selectedRange,
      yRangeMin,
      yRangeMax,
      liveText
    });
  };

  const handlePlay = () => {
    setPausedSnapshot(null);
  };

  const handleAutoScale = () => {
    const values = cleanedHistory
      .map((item) => ({
        time: new Date(item.timestamp).getTime(),
        value: Number(item[config.key])
      }))
      .filter(
        (item) =>
          !Number.isNaN(item.time) &&
          !Number.isNaN(item.value)
      );

    if (!values.length) return;

    const minValue = Math.min(...values.map((v) => v.value));
    const maxValue = Math.max(...values.map((v) => v.value));

    const minTime = Math.min(...values.map((v) => v.time));
    const maxTime = Math.max(...values.map((v) => v.time));

    const valueRange = Math.max(maxValue - minValue, config.yStep);
    const valuePadding = valueRange * 0.15;

    const nextMinY = clamp(
      Math.floor((minValue - valuePadding) / config.yStep) * config.yStep,
      config.yMin,
      config.yMax
    );

    const nextMaxY = clamp(
      Math.ceil((maxValue + valuePadding) / config.yStep) * config.yStep,
      config.yMin,
      config.yMax
    );

    setYRangeMin(Math.min(nextMinY, nextMaxY - config.yStep));
    setYRangeMax(Math.max(nextMaxY, nextMinY + config.yStep));

    const totalDomainMs =
      domainEnd.getTime() - domainStart.getTime();

    const paddedStart =
      minTime - (maxTime - minTime) * 0.08;

    const paddedEnd =
      maxTime + (maxTime - minTime) * 0.08;

    const startPercent =
      ((paddedStart - domainStart.getTime()) /
        totalDomainMs) *
      100;

    const endPercent =
      ((paddedEnd - domainStart.getTime()) /
        totalDomainMs) *
      100;

    setRangeStartPercent(
      clamp(startPercent, 0, 100)
    );

    setRangeEndPercent(
      clamp(endPercent, 0, 100)
    );
  };

  const handleManualTimeChange = (which, field, rawValue) => {
    const parts = which === "start" ? { ...manualStartParts } : { ...manualEndParts };
    parts[field] = ["day", "month", "year", "hour", "minute", "second"].includes(field) ? Number(rawValue) : rawValue;
    let nextMs = aestPartsToDate(parts).getTime();
    const minMs = domainStart.getTime();
    const maxMs = domainEnd.getTime();
    const otherMs = which === "start" ? selectedRange.end.getTime() : selectedRange.start.getTime();
    nextMs = clamp(nextMs, minMs, maxMs);
    if (which === "start") nextMs = Math.min(nextMs, otherMs);
    if (which === "end") nextMs = Math.max(nextMs, otherMs);
    const totalMs = Math.max(maxMs - minMs, 1);
    const nextPercent = ((nextMs - minMs) / totalMs) * 100;
    if (which === "start") setRangeStartPercent(Math.min(nextPercent, rangeEndPercent));
    else setRangeEndPercent(Math.max(nextPercent, rangeStartPercent));
  };

  const handleYManualChange = (which, rawValue) => {
    const value = clamp(Number(rawValue), config.yMin, config.yMax);
    if (which === "min") setYRangeMin(Math.min(value, yRangeMax));
    else setYRangeMax(Math.max(value, yRangeMin));
  };

  return (
    <section className={`data-card extra-card ${config.className} ${isPaused ? "paused-panel" : ""}`}>
      <div className="card-header">
        <h2>Live {titleText}</h2>

        <div className="panel-header-actions" style={{ display: "flex", alignItems: "center", gap: "10px" }}>
          <span className="mini-stat">{activeLiveText}</span>

          {isPaused ? (
            <button
              type="button"
              className="panel-play-button"
              onClick={handlePlay}
            >
              ▶ Play
            </button>
          ) : (
            <button
              type="button"
              className="panel-pause-button"
              onClick={handlePause}
            >
              ⏸ Pause
            </button>
          )}

          <button
            type="button"
            className="panel-auto-button"
            onClick={handleAutoScale}
            style={{
              padding: "6px 10px",
              borderRadius: "10px",
              background: "var(--soft-panel)",
              border: "1px solid var(--accent-blue)",
              color: "var(--accent-blue)",
              fontWeight: 700,
              cursor: "pointer"
            }}
          >
            Auto Scale
          </button>
        </div>
      </div>

      <div className="graph-area" ref={graphAreaRef}>
        <GraphSvg graphData={{ ...graphData, hover: hoverPoint?.graphId === graphData.graphId ? hoverPoint : null }} onHover={setHoverPoint} />
      </div>
      <div className="graph-slider-wrap">
        <div className="graph-slider-labels">
          <span>{rangeLabels.left}</span>
          <span>{rangeLabels.right}</span>
        </div>

        <DualRangeSlider
          min={0}
          max={100}
          step={0.1}
          startValue={rangeStartPercent}
          endValue={rangeEndPercent}
          onStartChange={(value) => {
            if (value < rangeEndPercent) setRangeStartPercent(value);
          }}
          onEndChange={(value) => {
            if (value > rangeStartPercent) setRangeEndPercent(value);
          }}
        />

        <div className="graph-manual-inputs-row">
          <div className="graph-manual-inline">
            <span className="graph-manual-inline-label">Start:</span>
            <TimePartsSelect
              prefix={`${config.key}-start`}
              parts={manualStartParts}
              years={yearOptions}
              onChange={handleManualTimeChange}
            />
          </div>

          <div className="graph-manual-inline">
            <span className="graph-manual-inline-label">End:</span>
            <TimePartsSelect
              prefix={`${config.key}-end`}
              parts={manualEndParts}
              years={yearOptions}
              onChange={handleManualTimeChange}
            />
          </div>
        </div>

        <div className="graph-slider-labels graph-slider-labels-spaced">
          <span>{config.yMin}</span>
          <span>{config.yMax}</span>
        </div>

        <DualRangeSlider
          min={config.yMin}
          max={config.yMax}
          step={config.yStep}
          startValue={yRangeMin}
          endValue={yRangeMax}
          onStartChange={(value) => setYRangeMin(Math.min(value, yRangeMax))}
          onEndChange={(value) => setYRangeMax(Math.max(value, yRangeMin))}
        />

        <div className="graph-manual-inputs-row">
          <div className="graph-manual-inline">
            <span className="graph-manual-inline-label">{axisTitle} Min:</span>
            <select
              className="graph-select"
              value={yRangeMin}
              onChange={(e) => handleYManualChange("min", e.target.value)}
            >
              {yOptions.map((value) => (
                <option key={`${config.key}-y-min-${value}`} value={value}>
                  {Number(value).toFixed(config.yDigits)}
                </option>
              ))}
            </select>
          </div>

          <div className="graph-manual-inline">
            <span className="graph-manual-inline-label">{axisTitle} Max:</span>
            <select
              className="graph-select"
              value={yRangeMax}
              onChange={(e) => handleYManualChange("max", e.target.value)}
            >
              {yOptions.map((value) => (
                <option key={`${config.key}-y-max-${value}`} value={value}>
                  {Number(value).toFixed(config.yDigits)}
                </option>
              ))}
            </select>
          </div>

          <div className="graph-manual-inline">
            <span className="graph-manual-inline-label">Unit:</span>
            <select
              className="graph-select"
              value={selectedUnit}
              onChange={(e) => onUnitChange(e.target.value)}
            >
              {UNIT_OPTIONS.map((unit) => (
                <option key={unit.label} value={unit.value}>
                  {unit.label}
                </option>
              ))}
            </select>
          </div>
        </div>
      </div>
    </section>
  );
}

function GraphSvg({ graphData, onHover }) {
  const {
    width,
    height,
    leftPad,
    rightPad,
    topPad,
    bottomPad,
    segments,
    areaSegments,
    xAxis,
    yAxis,
    gridLinesX,
    gridLinesY,
    points,
    lineWidth,
    axisFontSize,
    titleFontSize,
    hover,
    xTitle,
    yTitle,
    graphId
  } = graphData;

  const graphBottom = height - bottomPad;
  const graphRight = width - rightPad;
  const graphTop = topPad;
  const hoverPoint = hover;

  return (
    <svg className="graph-svg" viewBox={`0 0 ${width} ${height}`} preserveAspectRatio="none">
      {gridLinesY.map((y, index) => (
        <line
          key={`grid-y-${index}`}
          x1={leftPad}
          y1={y}
          x2={graphRight}
          y2={y}
          className="graph-grid-line"
        />
      ))}

      {gridLinesX.map((x, index) => (
        <line
          key={`grid-x-${index}`}
          x1={x}
          y1={graphTop}
          x2={x}
          y2={graphBottom}
          className="graph-grid-line vertical"
        />
      ))}

      <line x1={leftPad} y1={graphBottom} x2={graphRight} y2={graphBottom} className="graph-axis-line" />
      <line x1={leftPad} y1={graphTop} x2={leftPad} y2={graphBottom} className="graph-axis-line" />

      {areaSegments.map((path, index) => (
        <path key={`area-${index}`} d={path} className="graph-area-fill" />
      ))}

      {segments.map((path, index) => (
        <path
          key={`line-${index}`}
          d={path}
          className="graph-line-path"
          strokeWidth={lineWidth}
        />
      ))}

      {points.map((point, index) => {
        const prevPoint = points[index - 1];
        const nextPoint = points[index + 1];

        const leftBoundary = prevPoint
          ? (prevPoint.x + point.x) / 2
          : leftPad;

        const rightBoundary = nextPoint
          ? (point.x + nextPoint.x) / 2
          : graphRight;

        return (
          <rect
            key={`point-hover-zone-${index}`}
            x={leftBoundary}
            y={graphTop}
            width={Math.max(rightBoundary - leftBoundary, 1)}
            height={graphBottom - graphTop}
            className="graph-hover-target"
            onMouseEnter={() => onHover({ ...point, graphId })}
            onMouseMove={() => onHover({ ...point, graphId })}
            onMouseLeave={() => onHover(null)}
          />
        );
      })}

      {hoverPoint ? (
        <g>
          <line
            x1={hoverPoint.x}
            y1={graphTop}
            x2={hoverPoint.x}
            y2={graphBottom}
            className="graph-cursor-line"
          />

          <circle
            cx={hoverPoint.x}
            cy={hoverPoint.y}
            r={6}
            className="graph-hover-point"
          />

          <TooltipBubble
            x={hoverPoint.x}
            y={hoverPoint.y}
            width={width}
            labelValue={hoverPoint.value}
            labelTime={formatTimestamp(hoverPoint.timestamp)}
          />
        </g>
      ) : null}

      {xAxis.map((tick, index) => (
        <g key={`x-tick-${index}`}>
          <text
            x={tick.x}
            y={graphBottom + axisFontSize * 1.8}
            fontSize={axisFontSize}
            textAnchor="middle"
            className="graph-axis-label"
          >
            {tick.label}
          </text>
        </g>
      ))}

      {yAxis.map((tick, index) => (
        <g key={`y-tick-${index}`}>
          <text
            x={leftPad - 12}
            y={tick.y + axisFontSize * 0.35}
            fontSize={axisFontSize}
            textAnchor="end"
            className="graph-axis-label"
          >
            {tick.value}
          </text>
        </g>
      ))}

      <text
        x={(leftPad + graphRight) / 2}
        y={height - axisFontSize * 1.1}
        fontSize={titleFontSize}
        textAnchor="middle"
        className="graph-axis-label"
      >
        {xTitle}
      </text>

      <text
        x={axisFontSize * 1.35}
        y={(graphTop + graphBottom) / 2}
        fontSize={titleFontSize}
        textAnchor="middle"
        transform={`rotate(-90 ${axisFontSize * 1.35} ${(graphTop + graphBottom) / 2})`}
        className="graph-axis-label"
      >
        {yTitle}
      </text>
    </svg>
  );
}

function TooltipBubble({ x, y, width, labelValue, labelTime }) {
  const boxWidth = 220;
  const boxHeight = 62;
  const margin = 14;

  const tooltipX = clamp(x + 14, margin, width - boxWidth - margin);
  const tooltipY = Math.max(y - boxHeight - 12, margin);

  return (
    <g pointerEvents="none">
      <rect
        x={tooltipX}
        y={tooltipY}
        width={boxWidth}
        height={boxHeight}
        rx={12}
        className="graph-tooltip-box"
      />
      <text
        x={tooltipX + 12}
        y={tooltipY + 24}
        className="graph-tooltip-title"
      >
        Value: {Number(labelValue).toFixed(2)}
      </text>
      <text
        x={tooltipX + 12}
        y={tooltipY + 45}
        className="graph-tooltip-time"
      >
        {labelTime}
      </text>
    </g>
  );
}

function DualRangeSlider({
  min,
  max,
  step,
  startValue,
  endValue,
  onStartChange,
  onEndChange
}) {
  const range = max - min;
  const startPercent = ((startValue - min) / range) * 100;
  const endPercent = ((endValue - min) / range) * 100;

  return (
    <div className="dual-slider">
      <div className="dual-slider-track" />
      <div
        className="dual-slider-range"
        style={{
          left: `${startPercent}%`,
          width: `${endPercent - startPercent}%`
        }}
      />

      <input
        type="range"
        min={min}
        max={max}
        step={step}
        value={startValue}
        onChange={(e) => onStartChange(Number(e.target.value))}
        className="slider-thumb"
      />

      <input
        type="range"
        min={min}
        max={max}
        step={step}
        value={endValue}
        onChange={(e) => onEndChange(Number(e.target.value))}
        className="slider-thumb"
      />
    </div>
  );
}

function TimePartsSelect({ prefix, parts, years, onChange }) {
  return (
    <div className="graph-manual-selects">
      <select
        className="graph-select"
        value={parts.day}
        onChange={(e) => onChange(prefix, "day", e.target.value)}
      >
        {Array.from({ length: 31 }, (_, i) => (
          <option key={i} value={i + 1}>
            {String(i + 1).padStart(2, "0")}
          </option>
        ))}
      </select>

      <select
        className="graph-select"
        value={parts.month}
        onChange={(e) => onChange(prefix, "month", e.target.value)}
      >
        {MONTH_NAMES.map((m, i) => (
          <option key={m} value={i + 1}>
            {m}
          </option>
        ))}
      </select>

      <select
        className="graph-select graph-select-year"
        value={parts.year}
        onChange={(e) => onChange(prefix, "year", e.target.value)}
      >
        {years.map((y) => (
          <option key={y} value={y}>
            {y}
          </option>
        ))}
      </select>

      <select
        className="graph-select"
        value={parts.hour}
        onChange={(e) => onChange(prefix, "hour", e.target.value)}
      >
        {Array.from({ length: 12 }, (_, i) => (
          <option key={i} value={i + 1}>
            {String(i + 1).padStart(2, "0")}
          </option>
        ))}
      </select>

      <select
        className="graph-select"
        value={parts.minute}
        onChange={(e) => onChange(prefix, "minute", e.target.value)}
      >
        {Array.from({ length: 60 }, (_, i) => (
          <option key={i} value={i}>
            {String(i).padStart(2, "0")}
          </option>
        ))}
      </select>

      <select
        className="graph-select"
        value={parts.second}
        onChange={(e) => onChange(prefix, "second", e.target.value)}
      >
        {Array.from({ length: 60 }, (_, i) => (
          <option key={i} value={i}>
            {String(i).padStart(2, "0")}
          </option>
        ))}
      </select>

      <select
        className="graph-select graph-select-ampm"
        value={parts.ampm}
        onChange={(e) => onChange(prefix, "ampm", e.target.value)}
      >
        <option value="AM">AM</option>
        <option value="PM">PM</option>
      </select>
    </div>
  );
}

function SummaryBlock({ title, stats, digits = 2, suffix = "" }) {
  return (
    <div className="analysis-summary-block">
      <h3>{title}</h3>

      <div className="analysis-grid">
        <div className="analysis-row">
          <span className="label">Latest Value:</span>
          <span className="value">{formatSummaryValue(stats.current, digits, suffix)}</span>
        </div>

        <div className="analysis-row">
          <span className="label">Average:</span>
          <span className="value">{formatSummaryValue(stats.average, digits, suffix)}</span>
        </div>

        <div className="analysis-row">
          <span className="label">Min:</span>
          <span className="value">{formatSummaryValue(stats.min, digits, suffix)}</span>
        </div>

        <div className="analysis-row">
          <span className="label">Max:</span>
          <span className="value">{formatSummaryValue(stats.max, digits, suffix)}</span>
        </div>

        <div className="analysis-row">
          <span className="label">Trend:</span>
          <span className="value">{stats.trend}</span>
        </div>
      </div>
    </div>
  );
}

function getMetricStatsForSeries(series) {
  return Object.fromEntries(
    METRIC_GRAPH_CONFIGS.map((config) => [
      config.key,
      getSeriesStats(series, config.key)
    ])
  );
}

function PeriodAnalysisBlock({title, metricStats, getMetricAxisTitle, metricUnits}) {
  return (
    <div className="analysis-summary-block period-block">
      <h3>{title}</h3>

      {METRIC_GRAPH_CONFIGS.map((config) => {
        const stats = metricStats[config.key];

        return (
          <div key={`period-${title}-${config.key}`} className="period-metric-section">
            <div className="period-metric-heading">{getMetricAxisTitle(config)}:</div>

            <div className="analysis-grid">
              <div className="analysis-row">
                <span className="label">Average:</span>
                <span className="value">
                  {formatSummaryValue(stats.average, config.yDigits, metricUnits[config.key] ? ` ${metricUnits[config.key]}` : "")}
                </span>
              </div>

              <div className="analysis-row">
                <span className="label">Min:</span>
                <span className="value">
                  {formatSummaryValue(stats.min, config.yDigits, metricUnits[config.key] ? ` ${metricUnits[config.key]}` : "")}
                </span>
              </div>

              <div className="analysis-row">
                <span className="label">Max:</span>
                <span className="value">
                  {formatSummaryValue(stats.max, config.yDigits, metricUnits[config.key] ? ` ${metricUnits[config.key]}` : "")}
                </span>
              </div>

              <div className="analysis-row">
                <span className="label">Trend:</span>
                <span className="value">{stats.trend}</span>
              </div>
            </div>
          </div>
        );
      })}
    </div>
  );
}

function ThresholdSlider({
  label,
  min,
  max,
  step,
  value,
  onChange,
  displayValue
}) {
  return (
    <div className="threshold-control">
      <div className="threshold-header">
        <span className="label">{label}</span>
        <span className="value">{displayValue}</span>
      </div>

      <div className="threshold-slider-wrap">
        <div className="graph-slider-labels">
          <span>{min}</span>
          <span>{max}</span>
        </div>

        <input
          type="range"
          min={min}
          max={max}
          step={step}
          value={value}
          onChange={(e) => onChange(Number(e.target.value))}
          className="single-slider"
        />
      </div>
    </div>
  );
}


function inlineSvgStyles(sourceNode, clonedNode) {
  const styleProperties = [
    "fill",
    "stroke",
    "stroke-width",
    "stroke-linecap",
    "stroke-linejoin",
    "stroke-dasharray",
    "opacity",
    "font-family",
    "font-size",
    "font-weight",
    "text-anchor",
    "dominant-baseline"
  ];

  const sourceElements = [sourceNode, ...sourceNode.querySelectorAll("*")];
  const clonedElements = [clonedNode, ...clonedNode.querySelectorAll("*")];

  sourceElements.forEach((sourceElement, index) => {
    const clonedElement = clonedElements[index];
    if (!clonedElement) return;

    const computed = window.getComputedStyle(sourceElement);

    styleProperties.forEach((property) => {
      const value = computed.getPropertyValue(property);
      if (value) clonedElement.style.setProperty(property, value);
    });
  });
}

function setupGraphResizeObserver(ref, setSize) {
  if (!ref.current) return () => {};

  const updateSize = () => {
    const width = Math.max(ref.current?.clientWidth || 700, 520);
    const height = Math.max(Math.min(width / 2, 620), 320);
    setSize({ width, height });
  };

  updateSize();

  const observer = new ResizeObserver(updateSize);
  observer.observe(ref.current);

  return () => observer.disconnect();
}

function getGraphData({
  series,
  xStart,
  xEnd,
  yMin,
  yMax,
  width,
  height,
  valueKey = "temperature",
  xTitle,
  yTitle
}) {
  const axisFontSize = clamp(Math.min(width, height) * 0.024, 9, 13);
  const titleFontSize = clamp(Math.min(width, height) * 0.028, 11, 16);
  const leftPad = clamp(width * 0.09, 64, 92);
  const rightPad = clamp(width * 0.02, 18, 28);
  const topPad = clamp(height * 0.07, 18, 32);
  const bottomPad = clamp(height * 0.22, 76, 112);
  const plotWidth = width - leftPad - rightPad;
  const plotHeight = height - topPad - bottomPad;
  const lineWidth = clamp(width * 0.003, 2.5, 4.5);

  const startMs = new Date(xStart).getTime();
  const endMs = new Date(xEnd).getTime();
  const xRangeMs = Math.max(endMs - startMs, 1);
  const minValue = Math.min(yMin, yMax);
  const maxValue = Math.max(yMin, yMax === yMin ? yMin + 1 : yMax);
  const yRange = Math.max(maxValue - minValue, 1);

  const xToPoint = (timestamp) =>
    leftPad + ((new Date(timestamp).getTime() - startMs) / xRangeMs) * plotWidth;

  const yToPoint = (value) =>
    height - bottomPad - ((value - minValue) / yRange) * plotHeight;

  const sortedSeries = [...series].sort(
    (a, b) => new Date(a.timestamp) - new Date(b.timestamp)
  );

  const diffs = [];
  for (let i = 1; i < sortedSeries.length; i += 1) {
    const prev = new Date(sortedSeries[i - 1].timestamp).getTime();
    const curr = new Date(sortedSeries[i].timestamp).getTime();
    const diff = curr - prev;
    if (diff > 0) diffs.push(diff);
  }

  const medianDiff = diffs.length
    ? [...diffs].sort((a, b) => a - b)[Math.floor(diffs.length / 2)]
    : 5 * 60 * 1000;

  const gapThreshold = Math.max(medianDiff * 3, 30 * 60 * 1000);

  const pointSegments = [];
  let currentSegment = [];

  sortedSeries.forEach((item, index) => {
    const value = Number(item[valueKey]);
    if (Number.isNaN(value)) return;

    const point = {
      x: xToPoint(item.timestamp),
      y: yToPoint(value),
      value,
      timestamp: item.timestamp
    };

    if (!currentSegment.length) {
      currentSegment.push(point);
      return;
    }

    const prevTime = new Date(sortedSeries[index - 1].timestamp).getTime();
    const currTime = new Date(item.timestamp).getTime();

    if (currTime - prevTime > gapThreshold) {
      pointSegments.push(currentSegment);
      currentSegment = [point];
    } else {
      currentSegment.push(point);
    }
  });

  if (currentSegment.length) pointSegments.push(currentSegment);

  const segments = pointSegments
    .filter((segment) => segment.length)
    .map((segment) =>
      segment
        .map((point, index) => `${index === 0 ? "M" : "L"} ${point.x} ${point.y}`)
        .join(" ")
    );

  const areaSegments = pointSegments
    .filter((segment) => segment.length)
    .map((segment) => {
      const line = segment
        .map((point, index) => `${index === 0 ? "M" : "L"} ${point.x} ${point.y}`)
        .join(" ");
      const first = segment[0];
      const last = segment[segment.length - 1];
      return `${line} L ${last.x} ${height - bottomPad} L ${first.x} ${height - bottomPad} Z`;
    });

  const xAxis = buildXAxisTicks(startMs, endMs, width, leftPad, rightPad);
  const yAxis = buildYAxisTicks(minValue, maxValue, height, topPad, bottomPad, plotHeight);

  return {
    width,
    height,
    leftPad,
    rightPad,
    topPad,
    bottomPad,
    plotWidth,
    plotHeight,
    lineWidth,
    axisFontSize,
    titleFontSize,
    segments,
    areaSegments,
    xAxis,
    yAxis,
    gridLinesX: xAxis.map((tick) => tick.x),
    gridLinesY: yAxis.map((tick) => tick.y),
    points: pointSegments.flat(),
    series,
    xTitle,
    yTitle,
    graphId: `${valueKey}-${width}-${height}-${minValue}-${maxValue}`
  };
}

function buildXAxisTicks(startMs, endMs, width, leftPad, rightPad) {
  const plotWidth = width - leftPad - rightPad;
  const tickCount = clamp(Math.round(plotWidth / 170), 6, 11);

  return Array.from({ length: tickCount }, (_, i) => {
    const ratio = tickCount === 1 ? 0 : i / (tickCount - 1);
    const ms = startMs + ratio * (endMs - startMs);
    const x = leftPad + ratio * plotWidth;

    return {
      x,
      label: formatGraphTimeAEST(ms, tickCount > 8)
    };
  });
}

function buildYAxisTicks(minValue, maxValue, height, topPad, bottomPad, plotHeight) {
  const range = Math.max(maxValue - minValue, 1);
  const targetTicks = clamp(Math.round(plotHeight / 70), 4, 8);
  const step = getNiceStep(range / Math.max(targetTicks - 1, 1));
  const first = Math.ceil(minValue / step) * step;
  const ticks = [];

  for (let value = first; value <= maxValue + step * 0.25; value += step) {
    const y = height - bottomPad - ((value - minValue) / range) * plotHeight;
    const decimals = step < 1 ? 1 : 0;

    ticks.push({
      value: Number(value.toFixed(decimals)).toFixed(decimals),
      y
    });
  }

  if (!ticks.length) {
    return [
      { value: minValue.toFixed(1), y: height - bottomPad },
      { value: maxValue.toFixed(1), y: topPad }
    ];
  }

  return ticks;
}

function getNiceStep(rawStep) {
  const steps = [0.5, 1, 2, 5, 10, 20, 25, 50];
  return steps.find((step) => rawStep <= step) || 50;
}

function getSeriesStats(series, key) {
  const values = series
    .map((item) => Number(item[key]))
    .filter((value) => !Number.isNaN(value));

  if (!values.length) {
    return {
      current: null,
      average: null,
      min: null,
      max: null,
      trend: "Waiting"
    };
  }

  const current = values[values.length - 1];
  const average = values.reduce((sum, value) => sum + value, 0) / values.length;
  const min = Math.min(...values);
  const max = Math.max(...values);

  let trend = "Stable";

  if (values.length >= 2) {
    const first = values[0];
    const last = values[values.length - 1];
    const delta = last - first;
    const range = Math.max(max - min, 1);
    const ratio = Math.abs(delta) / range;

    if (ratio < 0.1) {
      trend = "Stable";
    } else if (delta > 0) {
      trend = "Rising";
    } else {
      trend = "Falling";
    }
  }

  return {
    current,
    average,
    min,
    max,
    trend
  };
}

function formatSummaryValue(value, digits = 2, suffix = "") {
  if (value == null || Number.isNaN(value)) return "Waiting";
  return `${Number(value).toFixed(digits)}${suffix}`;
}

function buildAlertEvents({ series, thresholds = {}, configs = METRIC_GRAPH_CONFIGS }) {
  const alerts = [];

  series.forEach((item) => {
    configs.forEach((config) => {
      const value = Number(item[config.key]);
      const threshold = Number(thresholds[config.key]);

      if (!Number.isNaN(value) && !Number.isNaN(threshold) && value > threshold) {
        alerts.push({
          type: config.key,
          message:
            `${config.alertTitle || config.yTitle} exceeded threshold at ` +
            `${formatTimestamp(item.timestamp)} ` +
            `(${value.toFixed(config.yDigits)}` +
            `${config.selectedUnit ? ` ${config.selectedUnit}` : ""} > ` +
            `${threshold.toFixed(config.yDigits)}` +
            `${config.selectedUnit ? ` ${config.selectedUnit}` : ""})`
        });
      }
    });
  });

  return alerts;
}

function getAestDateParts(dateInput) {
  const date = new Date(dateInput);
  const shifted = new Date(date.getTime() + AEST_OFFSET_MS);

  return {
    year: shifted.getUTCFullYear(),
    month: shifted.getUTCMonth(),
    day: shifted.getUTCDate(),
    hour: shifted.getUTCHours(),
    minute: shifted.getUTCMinutes(),
    second: shifted.getUTCSeconds()
  };
}

function getAestDayWindow(dayOffset = 0) {
  const now = new Date();
  const shiftedNow = new Date(now.getTime() + AEST_OFFSET_MS);

  const startUtcMs =
    Date.UTC(
      shiftedNow.getUTCFullYear(),
      shiftedNow.getUTCMonth(),
      shiftedNow.getUTCDate() - dayOffset,
      0,
      0,
      0
    ) - AEST_OFFSET_MS;

  const endUtcMs = startUtcMs + 24 * 60 * 60 * 1000;

  return {
    start: new Date(startUtcMs),
    end: new Date(endUtcMs)
  };
}

function filterSeriesByTimeWindow(series, dayOffset, startHour, endHour) {
  const dayWindow = getAestDayWindow(dayOffset);

  return series.filter((item) => {
    const t = new Date(item.timestamp).getTime();
    if (Number.isNaN(t)) return false;
    if (t < dayWindow.start.getTime() || t >= dayWindow.end.getTime()) return false;

    const aestParts = getAestDateParts(item.timestamp);
    const hour = aestParts.hour;

    return hour >= startHour && hour < endHour;
  });
}

function removeFirstPointOfEachSeries(series, gapThresholdMs = 5 * 60 * 1000) {
  if (!Array.isArray(series) || series.length === 0) return [];

  const sorted = [...series].sort(
    (a, b) => new Date(a.timestamp) - new Date(b.timestamp)
  );

  return sorted.filter((item, index) => {
    if (index === 0) return false;

    const prevTime = new Date(sorted[index - 1].timestamp).getTime();
    const currTime = new Date(item.timestamp).getTime();

    if (Number.isNaN(prevTime) || Number.isNaN(currTime)) return false;

    const gap = currTime - prevTime;

    if (gap > gapThresholdMs) {
      return false;
    }

    return true;
  });
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

function formatTimestamp(timestamp) {
  if (!timestamp || timestamp === "--") return "--";
  const date = new Date(timestamp);
  if (Number.isNaN(date.getTime())) return timestamp;

  return new Intl.DateTimeFormat("en-AU", {
    timeZone: "Australia/Brisbane",
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
    second: "2-digit",
    hour12: true
  }).format(date);
}

function formatGraphTimeAEST(dateInput, compact = false) {
  if (!dateInput) return "--";
  const date = new Date(dateInput);
  if (Number.isNaN(date.getTime())) return "--";

  return new Intl.DateTimeFormat("en-AU", {
    timeZone: "Australia/Brisbane",
    month: "short",
    day: "numeric",
    ...(compact
      ? { hour: "numeric", hour12: true }
      : { hour: "numeric", minute: "2-digit", hour12: true })
  }).format(date);
}

function toAestParts(dateInput) {
  const date = new Date(dateInput);
  const shifted = new Date(date.getTime() + AEST_OFFSET_MS);
  const hour24 = shifted.getUTCHours();

  return {
    year: shifted.getUTCFullYear(),
    month: shifted.getUTCMonth() + 1,
    day: shifted.getUTCDate(),
    hour: hour24 % 12 || 12,
    minute: shifted.getUTCMinutes(),
    second: shifted.getUTCSeconds(),
    ampm: hour24 >= 12 ? "PM" : "AM"
  };
}

function aestPartsToDate(parts) {
  let hour24 = Number(parts.hour) % 12;
  if (parts.ampm === "PM") hour24 += 12;

  return new Date(
    Date.UTC(
      Number(parts.year),
      Number(parts.month) - 1,
      Number(parts.day),
      hour24,
      Number(parts.minute),
      Number(parts.second)
    ) - AEST_OFFSET_MS
  );
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function isHexColor(value) {
  return typeof value === "string" && /^#([0-9A-Fa-f]{6})$/.test(value);
}

export default App;