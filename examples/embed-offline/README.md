# Offline embed (no build step)

The smallest possible integration of `@forevka/wordcanvas`: a static HTML page
that mounts the editor in fully offline mode (no backend, no sync, no share),
using a browser **import map** to resolve the package — no bundler, no tooling.

## Run it

From the **repo root** (so `/node_modules/...` in the import map resolves):

```sh
npm install            # once, links the workspace packages
npm run build:lib      # produces frontend/dist-lib that the import map points at
npx serve .            # or: python -m http.server
```

Then open `http://localhost:3000/examples/embed-offline/` (port depends on your
static server).

## What to look at

- `index.html` — the import map mapping `@forevka/wordcanvas` to the built bundle,
  plus a `#loader` overlay (a label + a progress bar).
- `app.js` — `import { WordCanvas } from "@forevka/wordcanvas"; new WordCanvas({ container })`.
- `app.js` also wires **`onLoadProgress`** to drive that overlay. On a cold load the
  editor JS chunk and the bundled fonts (~9 MB) stream before the editor is
  interactive; the callback reports an overall `percent` (0..1, monotonic) plus a
  coarse `phase` (`"bundle"` → `"fonts"` → `"ready"`), so the bar fills smoothly and
  the overlay fades out on `"ready"`. To see it on a slow connection, hard-reload
  with DevTools → Network throttled to *Slow 3G* and *Disable cache* checked.
- `app.js` also wires a **`resolveField`** callback. Open a `.docx` containing a
  developer-defined field — a paragraph whose `w:instrText` is something like
  ` MYCHART "sales-2026" ` — then right-click it and choose **Update Field
  (MYCHART)**. The engine calls `resolveField({ name, instruction })`, parses the
  OOXML it returns, and splices it in as the field's result. In production you'd
  forward that request to your backend (which renders the OOXML) instead of
  synthesizing it locally. Built-in `TOC` fields show **Update Field (TOC)** and
  regenerate locally (no backend).

## Debugging a document with an AI agent (WebMCP)

`app.js` sets **`agentTools: true`**, which exposes the live editor to AI agents
over [WebMCP](https://webmcp.dev) — the standard `navigator.modelContext` API.
This is the "a user reported a document that renders weirdly" workflow: open the
offending `.docx` locally (the file picker, top-right), connect an agent, and let
it inspect and fix the document.

**Connect an agent** with either:

- the **WebMCP browser extension** (it discovers the page's `navigator.modelContext`
  tools and bridges them to your MCP client / Claude Desktop), or
- **Chrome DevTools MCP** — point your agent at the tab.

The polyfill that installs `navigator.modelContext` is bundled into the editor and
loaded lazily only because `agentTools` is set, so embedders that don't opt in pay
nothing.

**Tools the agent gets** (all on by default; restrict with
`agentTools: { capabilities: ["read"] }`):

| Bucket | Tools |
| --- | --- |
| read & inspect | `get_document`, `get_selection`, `search_document`, `inspect_layout`, `get_document_stats` |
| suggest & comment | `set_mode`, `get_review`, `add_comment`, `reply_to_comment`, `resolve_thread`, `accept_suggestion`, `reject_suggestion` |
| direct edits | `replace_text`, `insert_text`, `format_text`, `set_alignment`, `select_range`, `undo`, `redo`, `set_document` |

`inspect_layout` is the debugging workhorse: it dumps the laid-out geometry
(per page → block → line → text-fragment positions, in page-local CSS px). Example
prompts:

- *"Call `inspect_layout` for page 1 and tell me why the heading overlaps the table."*
- *"Find the paragraph whose fragments have `whitespaceCollapsed: true` and show its text."*
- *"Switch to suggest mode and replace every 'colour' with 'color', then add a comment explaining why."*

To restrict the agent to read-only inspection, set:

```js
new WordCanvas({ container, agentTools: { capabilities: ["read"] } });
```

In a real app installed from npm, drop the import map and let your bundler resolve
the bare specifier:

```js
import { WordCanvas } from "@forevka/wordcanvas";
new WordCanvas({ container: document.getElementById("editor") });
```
