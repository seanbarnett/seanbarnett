# Cloudflare Worker — private dashboard hosting

Why this exists: GitHub Pages on this repo has no way to require a login
(that needs a paid GitHub plan), and browsers can't call the Google Apps
Script `/exec` URL directly — Google blocks both plain `fetch()` (no CORS
headers) and JSONP `<script>` loads (anti-abuse redirect) from third-party
pages. A Cloudflare Worker fixes both problems in one free service:

- It requires a username + password before serving anything, so the
  dashboard isn't reachable by "anyone who has the URL."
- It fetches the Apps Script data **server-side** (`/data` route) — that's
  a plain server-to-server request with no browser involved, so none of
  Google's browser-side CORS/anti-abuse restrictions apply.

## One-time setup (no CLI required)

1. Create a free account at [dash.cloudflare.com](https://dash.cloudflare.com)
   if you don't have one.
2. Go to **Workers & Pages → Create → Create Worker**. Give it any name
   (e.g. `barnett-dashboard`) and click **Deploy** to scaffold it.
3. Click **Edit code** and replace the entire default script with the full
   contents of `worker.js` in this folder.
4. Click **Deploy**.
5. Go to the Worker's **Settings → Variables and Secrets** and add three:
   | Name | Value | Type |
   |---|---|---|
   | `AUTH_USER` | a username you choose, e.g. `sean` | Plaintext |
   | `AUTH_PASS` | a real password (not a Google password) | **Encrypt** |
   | `GAS_URL` | your Apps Script `/exec` URL | Plaintext |
6. Save — this redeploys the Worker with those values available as `env.*`.
7. Visit `https://<worker-name>.<your-subdomain>.workers.dev`. Your browser
   will prompt for the username/password from step 5 before showing
   anything.

Share that URL + the username/password with Cecilia (not the Apps Script
URL itself — the Worker is the front door now).

## Re-deploying after editing `dashboard.html`

The dashboard's HTML is embedded directly inside `worker.js` as a string
(Cloudflare's dashboard editor only accepts a single file, so there's no
separate static-asset upload here). Whenever `dashboard.html` changes:

```bash
python3 cloudflare-worker/build.py
```

This regenerates `worker.js` from the current `dashboard.html`. Copy the
new `worker.js` contents into the Cloudflare dashboard's **Edit code**
view and click **Deploy** again.

## Notes

- The `gas/Code.gs` Apps Script deployment must still be set to
  **Execute as: Me / Who has access: Anyone** — "Anyone" here just means
  "reachable if you know the exact random `/exec` URL," and now only this
  Worker knows it (it's never sent to any browser).
- Do **not** commit real values for `AUTH_USER`/`AUTH_PASS` anywhere in
  this repo — they live only in the Worker's encrypted environment
  variables, set through the Cloudflare dashboard.
