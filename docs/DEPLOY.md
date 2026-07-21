# Deploying the full stack

The whole app (editor + examples, the Node backend, Postgres) deploys to a single
VPS as Docker containers, served by an in-container Caddy that the **host** Caddy
fronts for TLS. One script does it: [`deploy.ps1`](./deploy.ps1).

## Usage

```powershell
.\deploy.ps1 -RemoteHost root@your-server -Domain doc-editor.example.com
```

`-RemoteHost` and `-Domain` can also come from `$env:CW_DEPLOY_HOST` /
`$env:CW_DEPLOY_DOMAIN` (so no host is hard-coded in this public repo). Optional:
`-RemotePath` (default `/opt/canvas-word`), `-EdgePort` (host port the web edge
binds on `127.0.0.1`, default `3003`).

## What it does

1. **Pack** the working tree with `tar` (excludes `node_modules`, `dist*`,
   `.git`, `.env`, `local-db-data`, fixtures/`*.docx`, etc.).
2. **Upload** the tarball over `scp`.
3. **Extract + build** on the VPS: untar (preserving `.env` and `local-db-data`),
   seed `.env` with a generated DB password + admin login on first run, then
   `docker compose --profile app up -d --build`. Migrations run via the `flyway`
   service against the `postgres` service before `backend`/`web` start.
4. **Smoke test** `https://<domain>/openapi.json` (expects `200`).

> The script deploys the **local working tree**, not a git ref — commit (and ideally
> push) before deploying so what's live matches origin.

## Caching / cache-busting (important)

The in-container Caddy ([`web/Caddyfile`](./web/Caddyfile)) sets an explicit cache
policy so a new deployment is picked up **without a hard reload**:

- **HTML entry documents and SPA fallbacks → `Cache-Control: no-cache`.** Always
  revalidated, so each deploy's new content-hashed asset references are seen on the
  next normal load.
- **Content-hashed bundles under `assets/` → `Cache-Control: public, max-age=31536000, immutable`.**
  Safe because Vite fingerprints every JS/CSS/**worker** chunk — the filename
  changes on every rebuild — so old chunks (including the export/import web workers)
  are never requested again, and the multi-MB worker bundle isn't re-fetched on
  repeat visits.
- Backend paths (`/docs`, `/media`, `/ws`, `/admin`, `/upload`, `/openapi.json`,
  `/swagger`) are left to the backend's own headers.

Without this, `file_server` sent no `Cache-Control`, browsers heuristically cached
`index.html`, and after a deploy they kept loading the old hashed chunks + worker
from cache (running stale worker code) until a hard reload.

**One-time caveat:** clients that cached the old `index.html` *before* this policy
shipped may need one more load (until their heuristic-freshness window lapses) to
switch over. Every deploy after that invalidates cleanly.

After changing `web/Caddyfile`, validate before deploying:

```powershell
docker run --rm -v "${PWD}/web/Caddyfile:/etc/caddy/Caddyfile:ro" `
  caddy:2-alpine caddy validate --config /etc/caddy/Caddyfile --adapter caddyfile
```
