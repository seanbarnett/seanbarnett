#!/usr/bin/env python3
"""
Regenerates cloudflare-worker/worker.js by embedding the current
dashboard.html as a JS string literal inside the Worker template below.

Run this any time you edit dashboard.html, then re-paste the resulting
worker.js into the Cloudflare dashboard (or `wrangler deploy` if you use
the CLI).

Usage:
    python3 cloudflare-worker/build.py
"""
import json
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
DASHBOARD_HTML_PATH = ROOT / 'dashboard.html'
OUTPUT_PATH = pathlib.Path(__file__).resolve().parent / 'worker.js'

WORKER_TEMPLATE = '''// Cloudflare Worker — private front door for the financial dashboard.
//
// What this does:
//   1. Requires HTTP Basic Auth before serving anything (401 challenge
//      otherwise) — keeps the dashboard off the public internet even though
//      Cloudflare Workers themselves don't have a private-visibility toggle.
//   2. Serves the dashboard HTML at "/".
//   3. Proxies "/data" server-side to the real Google Apps Script /exec URL.
//      This request never touches a browser, so none of the CORS /
//      anti-abuse redirect issues Google's endpoint has with browser-side
//      fetch()/JSONP apply here — it's a plain server-to-server GET.
//
// SETUP (Cloudflare dashboard, no CLI needed):
//   1. workers.cloudflare.com -> Create Worker -> paste this whole file.
//   2. Settings -> Variables and Secrets -> add:
//        AUTH_USER   (plain text)  e.g. "sean"
//        AUTH_PASS   (encrypt!)    a real password, not a Google password
//        GAS_URL     (plain text)  your Apps Script /exec URL
//   3. Deploy. Your dashboard is now at https://<worker-name>.<subdomain>.workers.dev
//      and will prompt for the username/password you set above.
//
// Re-deploying after editing dashboard.html: re-run build.py to regenerate
// this file's DASHBOARD_HTML constant, then paste the updated worker.js
// back into the Cloudflare dashboard.

const DASHBOARD_HTML = __HTML_LITERAL__;

export default {
  async fetch(request, env) {
    const authResponse = checkAuth(request, env);
    if (authResponse) return authResponse;

    const url = new URL(request.url);

    if (url.pathname === '/data') {
      return proxyData(env);
    }

    if (url.pathname === '/' || url.pathname === '/dashboard.html') {
      return new Response(DASHBOARD_HTML, {
        headers: { 'content-type': 'text/html; charset=UTF-8' },
      });
    }

    return new Response('Not found', { status: 404 });
  },
};

function checkAuth(request, env) {
  const unauthorized = () =>
    new Response('Authentication required', {
      status: 401,
      headers: { 'WWW-Authenticate': 'Basic realm="Financial Dashboard", charset="UTF-8"' },
    });

  const header = request.headers.get('Authorization');
  if (!header || !header.startsWith('Basic ')) return unauthorized();

  let decoded;
  try {
    decoded = atob(header.slice('Basic '.length));
  } catch (e) {
    return unauthorized();
  }

  const sepIdx = decoded.indexOf(':');
  if (sepIdx === -1) return unauthorized();

  const user = decoded.slice(0, sepIdx);
  const pass = decoded.slice(sepIdx + 1);

  if (!timingSafeEqual(user, env.AUTH_USER) || !timingSafeEqual(pass, env.AUTH_PASS)) {
    return unauthorized();
  }

  return null; // authorized
}

function timingSafeEqual(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string') return false;
  const enc = new TextEncoder();
  const aBytes = enc.encode(a);
  const bBytes = enc.encode(b);
  if (aBytes.length !== bBytes.length) return false;
  let diff = 0;
  for (let i = 0; i < aBytes.length; i++) diff |= aBytes[i] ^ bBytes[i];
  return diff === 0;
}

async function proxyData(env) {
  try {
    const res = await fetch(env.GAS_URL, { redirect: 'follow' });
    const body = await res.text();
    return new Response(body, {
      status: res.status,
      headers: { 'content-type': 'application/json' },
    });
  } catch (e) {
    return new Response(JSON.stringify({ error: 'Proxy fetch failed: ' + e.message }), {
      status: 502,
      headers: { 'content-type': 'application/json' },
    });
  }
}
'''

def main():
    html = DASHBOARD_HTML_PATH.read_text(encoding='utf-8')
    html_literal = json.dumps(html)
    worker_js = WORKER_TEMPLATE.replace('__HTML_LITERAL__', html_literal)
    OUTPUT_PATH.write_text(worker_js, encoding='utf-8')
    print(f'Wrote {OUTPUT_PATH} ({len(worker_js)} bytes)')

if __name__ == '__main__':
    main()
