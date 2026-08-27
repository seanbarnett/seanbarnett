# Google Apps Script — Tiller Dashboard API

## Deploy (one-time setup)

1. Open your Tiller Google Sheet
2. Go to **Extensions → Apps Script**
3. Delete any existing code and paste the contents of `Code.gs`
4. Click **Deploy → New deployment**
   - Type: **Web App**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Click **Deploy** → authorize when prompted
6. Copy the `/exec` URL (looks like `https://script.google.com/macros/s/AKfy.../exec`)

## Wire up the dashboard

Open `dashboard.html` and set:

```js
const CONFIG = {
  GAS_URL: 'https://script.google.com/macros/s/AKfy.../exec',
  ...
};
```

The dashboard will now fetch live data every 5 minutes and fall back to static data on error.

## Re-deploying after code changes

If you edit `Code.gs` later, you must create a **new deployment** (or manage deployments → update an existing one) for changes to take effect. The `/exec` URL stays the same if you update an existing deployment.

## Security

- The Apps Script runs as **you**, so it has full access to the sheet
- The `/exec` endpoint is unauthenticated — only share the URL with people you trust
- No credentials are embedded in `dashboard.html`
