# Live collaboration embed (Vite)

[![Open in StackBlitz](https://developer.stackblitz.com/img/open_in_stackblitz.svg)](https://stackblitz.com/github/Forevka/canvas-word/tree/main/examples/embed-live)

The online variant of the embed, migrated from the old frontend `/live` page. It's
an ordinary **Vite** app that depends on `@forevka/wordcanvas` and consumes it by
package name — the realistic way an app integrates the editor. With a `backendUrl`
set, opening a document auto-publishes it, edits sync live, presence shows, and a
share link is produced.

Notice `vite.config.ts` needs **no special config** — no node polyfills, no
aliases. The package bundle is self-contained and ships its own workers.

## Run it

From the **repo root**:

```sh
npm install                 # once, links the workspace packages
npm run build:lib           # build @forevka/wordcanvas (the example imports the built bundle)
npm run db:up               # start Postgres + the collaboration backend (docker compose)
npm run dev:example         # start this example's Vite dev server (http://localhost:5180)
```

Open `http://localhost:5180`, enter a name, and edit. Click Share to publish and
get a link; open that link (it carries `?collab=<docId>`) in another tab/browser
to join the same document and see live collaboration + presence.

> The example imports the **built** `dist-lib`, so re-run `npm run build:lib`
> after changing the editor source.

## What to look at

- `src/main.ts` — `import { WordCanvas } from "@forevka/wordcanvas"`, mount with
  `backendUrl` + `user`, join via `?collab`. It also wires **`onLoadProgress`** to a
  `#loader` overlay (typed via the exported `LoadProgress`): the callback reports an
  overall `percent` (0..1) across the cold-load of the editor chunk + bundled fonts,
  so the bar fills and the overlay fades out on `phase: "ready"`.
- `src/identity.ts` — a demo-only identity prompt. The package ships none on
  purpose: a real embedder passes `user` directly (it owns auth).
- `vite.config.ts` — minimal; `optimizeDeps.exclude` serves the pre-built,
  worker-using bundle as-is in dev.

## Connect an AI agent (WebMCP)

`src/main.ts` sets **`agentTools: true`**, exposing this live document to AI agents
over [WebMCP](https://webmcp.dev) (the standard `navigator.modelContext` API) —
the "connect an agent as reviewer/editor to a shared document" workflow. Open the
page, connect an agent via the **WebMCP browser extension** or **Chrome DevTools
MCP**, then have it read, comment, suggest tracked changes, or edit directly. The
WebMCP polyfill is lazy-loaded only because `agentTools` is set. Restrict it with
`agentTools: { capabilities: ["read", "suggest"] }`.

## Deploying as `/live`

`base: "./"` makes the build path-relative, so `npm run build --workspace
@cw/example-embed-live` produces a `dist/` you can drop at `/srv/live` behind the
Caddy `@live` route (see `web/Caddyfile`). This is the page the editor's share
links point back to.
