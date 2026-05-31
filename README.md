# IAP Cloudflare Worker Setup

This project uses React + Vite for the frontend and a Cloudflare Worker for the backend. The backend serves these routes on the same domain:

- `/api/health`
- `/api/test-connection`
- `/api/device-data`

The Worker also serves the built frontend from the `dist` folder.

## 1. Open the project in VS Code

Open the folder that contains:

- `package.json`
- `wrangler.jsonc`
- `src`
- `worker`

## 2. Install Node packages

Open a terminal in VS Code and run:

```bash
npm install
```

## 3. Sign in to Cloudflare from VS Code

```bash
npx wrangler login
```

A browser window will open. Sign in to the Cloudflare account that owns your domain.

## 4. Set your TTN values in `wrangler.jsonc`

Open `wrangler.jsonc` and replace these placeholders:

- `TTN_REGION`
- `TTN_APP_ID`
- `TTN_DEVICE_ID`

Example:

```json
"vars": {
  "TTN_REGION": "au1",
  "TTN_APP_ID": "my-application-id",
  "TTN_DEVICE_ID": "my-device-id"
}
```

## 5. Add your TTN API key as a Cloudflare secret

Run:

```bash
npx wrangler secret put TTN_API_KEY
```

Paste your TTN API key when Wrangler asks for it.

## 6. Run the project locally

```bash
npm run dev
```

Wrangler will give you a local address. Open it in your browser.

Test these URLs locally:

- `/api/health`
- `/api/test-connection`
- `/api/device-data`

## 7. Build the frontend

```bash
npm run build
```

This creates the `dist` folder.

## 8. Deploy from VS Code with Wrangler

```bash
npm run deploy
```

That command builds the frontend and deploys the Worker + static assets together.

## 9. Attach your custom domain in Cloudflare

In Cloudflare:

1. Go to `Workers & Pages`
2. Open your Worker
3. Open `Domains & Routes`
4. Add your custom domain

If your domain is already attached to an old Pages project, remove it there first, then attach it to this Worker.

## 10. Test the live site

After deployment, test:

- `https://your-domain.com/`
- `https://your-domain.com/api/health`
- `https://your-domain.com/api/test-connection`
- `https://your-domain.com/api/device-data`

## Notes

- `TTN_API_KEY` is stored as a secret in Cloudflare, not in code.
- The frontend uses `/api/device-data`, so it works on the same domain after deployment.
- The TTN API key must have permission to read application traffic.
- If `api/test-connection` fails, check the TTN region, app ID, device ID, and API key.
