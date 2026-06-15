# WebMCP — AI agent tools

The editor can expose itself to AI agents over **[WebMCP](https://webmcp.dev)** —
the standard `navigator.modelContext` browser API for the Model Context Protocol.
With one option, a live editor instance publishes a set of **tools** an agent can
call to read, inspect, comment on, suggest changes to, and directly edit the open
document — in the same tab the user is looking at.

Two use cases drove this:

1. **Debug a document that "renders weirdly."** A user reports a document that
   looks wrong (overlapping text, bad page breaks, mis-placed table). Open the
   `.docx` locally and let an agent call `inspect_layout` to read the actual laid-out
   geometry (page → line → text-fragment positions, in CSS px) and pinpoint the
   problem.
2. **Connect an agent as a reviewer/editor.** Attach any MCP-capable agent to a
   specific document to read it, leave comments, propose tracked changes, or edit
   it directly.

---

## Quick start

```ts
import { WordCanvas } from "@forevka/wordcanvas";

new WordCanvas({
  container: document.getElementById("editor")!,
  agentTools: true, // ← expose the editor to agents over WebMCP
});
```

That's it. When `agentTools` is set, the editor lazy-loads the WebMCP polyfill,
installs `navigator.modelContext`, and registers the tool set against the live
editor. Then **connect an agent** (see [Connecting an agent](#connecting-an-agent)).

Restrict what agents can do with the object form:

```ts
new WordCanvas({ container, agentTools: { capabilities: ["read", "suggest"] } });
```

| `agentTools` value | Effect |
|---|---|
| `false` / omitted | No agent tooling. The WebMCP polyfill is **never loaded** — zero cost. |
| `true` | Register all tools (read & inspect + suggest & comment + direct edits). |
| `{ capabilities }` | Register only the listed buckets (`"read"`, `"suggest"`, `"edit"`). |
| `{ name }` | Prefix every tool name (e.g. `"doc1"` → `doc1_get_document`) — see [Multiple editors](#multiple-editors-on-one-page). |

`AgentToolsOptions` is part of the public type surface
(`frontend/types/wordcanvas.d.ts`).

---

## What WebMCP is (and which one this uses)

WebMCP lets a web page offer MCP tools directly to a browser-side agent, instead of
the agent driving the UI by screen-scraping/clicking. There are two ecosystems that
both go by "WebMCP":

- **The standards-track `navigator.modelContext` API** — the W3C Web Model Context
  surface. The **WebMCP browser extension** and **Chrome DevTools MCP** speak it,
  and it is the direction the platform is standardizing on.
- **The webmcp.dev `<script>` widget** (`@jason.today/webmcp`) — a simpler library
  with its own connect widget.

**This integration targets the standards `navigator.modelContext` API**, polyfilled
by [`@mcp-b/global`](https://www.npmjs.com/package/@mcp-b/global). Agents connect
through the WebMCP extension / Chrome DevTools MCP — **not** the webmcp.dev widget.

---

## Connecting an agent

1. Open a page running an editor with `agentTools` enabled (any of the bundled
   `examples/` work — see [Trying it in the examples](#trying-it-in-the-examples)).
2. Connect a WebMCP-capable client to that tab, either:
   - the **WebMCP browser extension** — it discovers the page's
     `navigator.modelContext` tools and bridges them to your MCP client (e.g. Claude
     Desktop), or
   - **Chrome DevTools MCP** — point your agent at the tab.
3. The agent now sees the editor's tools (`get_document`, `inspect_layout`,
   `replace_text`, …) and can call them. Changes appear live in the editor.

The polyfill also sets up the page-side transport the extension connects through, so
no extra wiring is required beyond `agentTools`.

---

## Tool reference

Tools are grouped into three capability buckets. `read` is always registered; the
others depend on `capabilities`.

### Read & inspect (`read` — always on)

| Tool | Arguments | Returns |
|---|---|---|
| `get_document` | `format?: "text" \| "json"` (default `text`) | Plain text (one line per paragraph) or the full document model as JSON. |
| `get_selection` | — | The current `{ anchor, focus }` selection and its selected text (`null` if no selection). |
| `search_document` | `query` (required), `matchCase?`, `wholeWord?` | `{ total, current }`. Highlights matches and moves the selection to the first one. |
| `inspect_layout` | `page?` (0-based), `blockId?`, `includeText?` (default `true`), `maxFragmentsPerLine?` | The laid-out geometry as JSON — see [Debugging rendering issues](#debugging-rendering-issues-inspect_layout). |
| `get_document_stats` | — | `{ pageCount, currentPage, blockCount, paragraphCount, mode, docId }`. |

### Suggest & comment (`suggest`)

| Tool | Arguments | Notes |
|---|---|---|
| `set_mode` | `mode: "edit" \| "suggest" \| "view"` (required) | Switch editor mode. In `suggest` mode, subsequent edits become tracked changes. Returns `{ ok, mode }` (`ok:false` if the mode isn't allowed). |
| `get_review` | — | The review overlay: tracked-change suggestions + comment threads. |
| `add_comment` | `body` (required), `find?` | Anchors a comment to a range. With `find`, it locates+selects that text first; otherwise comments on the current selection. Returns `{ threadId }`. |
| `reply_to_comment` | `threadId` (required), `body` (required) | Reply to an existing thread. |
| `resolve_thread` | `threadId` (required), `resolved?` (default `true`) | Resolve or reopen a thread. |
| `accept_suggestion` | `id?` | Accept one suggestion, or all when `id` is omitted. |
| `reject_suggestion` | `id?` | Reject one suggestion, or all when `id` is omitted. |

To produce a **tracked change**: `set_mode("suggest")`, then run an edit tool.

### Direct edits (`edit`)

| Tool | Arguments | Notes |
|---|---|---|
| `replace_text` | `find` (required), `replaceWith` (required), `all?`, `matchCase?`, `wholeWord?` | Find and replace. `all:true` replaces every match (returns `{ replaced }`); otherwise the first. Errors if not found. |
| `insert_text` | `text` (required), `find?` | Inserts at the current selection (replacing it if it's a range). With `find`, selects that text first — note this **replaces** the found text; use `replace_text` for plain replacement. |
| `format_text` | `find?`, `bold?`, `italic?`, `underline?`, `strikethrough?`, `color?`, `fontFamily?`, `fontSizePx?`, `highlightColor?`, `clear?` | Applies a character-style patch to the found/selected range. `clear:true` resets bold/italic/underline/strikethrough/highlight. |
| `set_alignment` | `align: "left" \| "center" \| "right" \| "justify"` (required) | Paragraph alignment of the selection. |
| `select_range` | `anchorBlockId`, `anchorOffset`, `focusBlockId`, `focusOffset` (all required) | Sets an explicit selection by model position. Discover block ids via `get_document(json)` / `inspect_layout`. |
| `undo` / `redo` | — | Undo / redo the last edit. |
| `set_document` | `json` (required) | Replace the whole document with a JSON model (same shape as `get_document(json)`). Drops undo history — use sparingly. |

**Addressing model.** Most edit tools are **text-anchored** (`find`) rather than
position-based: they reuse the editor's find/replace engine to locate ranges, so an
agent never has to reason about opaque block ids or UTF-16 offsets. For precise
control, `select_range` + the position-addressed tools are available.

**Result shape.** Every tool returns the MCP content shape
`{ content: [{ type: "text", text }], isError? }`. Failures (text not found, no
selection, invalid JSON) come back with `isError: true` and a message instead of
throwing.

---

## Debugging rendering issues (`inspect_layout`)

`inspect_layout` is the workhorse for "this document renders weirdly" reports. It
serializes the live **layout tree** — the absolutely-positioned geometry the paint
layer draws — into compact JSON. Coordinates are page-local CSS px (the same frame
the canvas draws in), rounded to 2 decimals.

Scope the output to keep it small:

- `inspect_layout()` — every page.
- `inspect_layout({ page: 1 })` — just page index 1 (0-based).
- `inspect_layout({ blockId: "p_3f2a" })` — just one block, searched on every page
  (including inside table cells).
- `inspect_layout({ includeText: false })` — geometry only, no rendered text.

Shape (abridged):

```jsonc
{
  "pageCount": 2,
  "defaultPageWidthPx": 816,
  "defaultPageHeightPx": 1056,
  "pages": [
    {
      "index": 0, "number": 1,
      "widthPx": 816, "heightPx": 1056,
      "marginPx": { "top": 96, "right": 96, "bottom": 96, "left": 96 },
      "contentTopPx": 96, "contentBottomPx": 960,
      "blockCount": 12, "headerBlocks": 0, "footerBlocks": 1,
      "blocks": [
        {
          "blockId": "p_1a2b", "kind": "paragraph", "x": 96, "y": 96,
          "firstLineIndex": 0,
          "lines": [
            {
              "y": 0, "height": 24, "ascent": 19,
              "fragments": [
                {
                  "x": 96, "width": 412,
                  "startOffset": 0, "endOffset": 41,
                  "style": { "font": "Calibri", "sizePx": 16, "bold": true },
                  "text": "The quick brown fox…",
                  "whitespaceCollapsed": true   // ← flagged when present
                }
              ]
            }
          ]
        }
        // images report { image: { width, height, srcKind, srcLength, clipped? } }
        // tables report { table: { x, y, width, height, colWidths, rows[]:{cells[]:{x,y,width,height,blocks[]}} } }
      ]
    }
  ]
}
```

What to look for when diagnosing a report:

- **Overlap / wrong stacking** — compare each `LineBox.y` (relative to its block's
  `y`) and `height`; compare block `x/y` against the page `marginPx`,
  `contentTopPx`/`contentBottomPx`, and table cell rects.
- **Spacing / "backspace deleted the wrong char"** — `whitespaceCollapsed: true`
  marks fragments where pretext collapsed whitespace (rendered text shorter than the
  model range). A frequent culprit in placement bugs.
- **Wrong font / size** — read each fragment's `style.font` / `style.sizePx`.
- **Bad page breaks** — which block lands on which page `index`, and `firstLineIndex`
  for paragraphs that split across pages.
- **Justification** — `wordSpacingPx` on fragments of justified lines.

A typical session: `get_document_stats` → `inspect_layout({ page: N })` → narrow to
a `blockId` → fix with `replace_text` / `format_text`, or report the structural bug.

---

## Multiple editors on one page

Each `WordCanvas` registers tools under the same global `navigator.modelContext`, so
several editors on one page would collide on tool names. Use **`name`** to namespace
each instance:

```ts
const a = new WordCanvas({ container: paneA, agentTools: { name: "editor0" } });
const b = new WordCanvas({ container: paneB, agentTools: { name: "editor1" } });
// → editor0_get_document, editor1_get_document, …
```

`destroy()` unregisters that instance's tools (via the shared `AbortSignal`), so ids
never need to be reused. See `examples/embed-multi`.

---

## Trying it in the examples

All bundled examples enable `agentTools` so they're testable after a deploy:

| Example | Setting | Best for |
|---|---|---|
| `embed-offline` | `agentTools: true` | Local debugging — open a reported `.docx` (file picker, top-right) and `inspect_layout` it. |
| `embed-live` | `agentTools: true` | Agent-as-reviewer/editor on a live, collaborative document. |
| `embed-multi` | `agentTools: { name: "editor<N>" }` | Targeting one editor among several (namespacing). |
| `playground` | `agentTools: true` | Inspecting the builder's output geometry while iterating (read & inspect). |

Run any of them (from the repo root), then connect an agent:

```sh
npm install
npm run build:lib            # build @forevka/wordcanvas (examples import the built bundle)
npm run dev:example          # embed-live   (or dev:multi / dev:playground)
# embed-offline: npx serve .  then open /examples/embed-offline/
```

> The examples import the **built** `dist-lib`, so re-run `npm run build:lib` after
> changing the editor source — the agent chunks ship inside it.

Example prompts once an agent is connected:

- *"Call `inspect_layout` for page 1 and tell me why the heading overlaps the table."*
- *"Find the paragraph whose fragments have `whitespaceCollapsed: true` and show its text."*
- *"Switch to suggest mode, replace every 'colour' with 'color', then add a comment explaining why."*

---

## Architecture & internals

- **Tool module** — `frontend/src/agent/webmcp.ts`. `registerAgentTools(editor, ctx,
  config)` registers the tools against the rich internal `Editor` and returns a
  disposer. The model-context object is injectable (`config.modelContext`) so the
  module is testable without a browser.
- **Layout serializer** — `frontend/src/agent/layoutDump.ts`. Pure
  `LayoutTree → JSON` used by `inspect_layout`.
- **Registration site** — `frontend/src/editorApp.ts`, right after the editor
  `handle` is built (where the full internal `Editor` is in scope). The public
  `EditorHandle` / `WordCanvas` API is **not** widened.
- **Editor accessors** — two read methods were added to the `Editor` interface in
  `frontend/src/index.ts` for the tools: `getLayoutTree()` and `setSelection()`.
- **Option plumbing** — `agentTools` flows `WordCanvasOptions` (`wordcanvas.ts`) →
  `WordCanvasRuntime` (`app/runtime.ts`) → registration in `editorApp.ts`. Public
  types in `frontend/types/wordcanvas.d.ts`.
- **Lifecycle** — all tools register under one `AbortController`; `WordCanvas`
  `destroy()` aborts it, unregistering every tool.

### Packaging & cost

- The polyfill ([`@mcp-b/global`](https://www.npmjs.com/package/@mcp-b/global), plus
  its `zod` / `zod-to-json-schema` peers) is a **devDependency, bundled, and
  lazy-imported** only inside `if (runtime.agentTools)`. The package's runtime
  `dependencies` stays empty — the **zero-runtime-dependency** promise holds.
- The agent code lands in separate chunks (`webmcp-*.js` ≈ 18 kB + a polyfill chunk
  ≈ 400 kB), fetched **only when `agentTools` is set**. Embedders who don't opt in
  download neither, and the editor's normal weight is unchanged.

### Tests

`frontend/src/agent/webmcp.test.ts` (vitest, Node — no DOM) covers tool wiring with
a stub `Editor` + fake model context: capability gating, name namespacing, text/json
reads, the `inspect_layout` dump (incl. block scoping), find-anchored
replace/format/insert, comment anchoring, the `set_document` host hook, and the
`AbortSignal` disposer.

---

## Security & operational notes

- **Tools act with the user's authority.** A connected agent can read the document
  and (with `edit`/`suggest`) modify it. Only enable `agentTools` where you intend an
  agent to have that access, and prefer the narrowest `capabilities` for the use case
  (e.g. `["read"]` for read-only inspection, `["read", "suggest"]` for a reviewer
  that proposes but never commits).
- **An agent still has to connect.** Tools are only reachable by a WebMCP client
  (extension / DevTools MCP) attached to that tab — they are not exposed to arbitrary
  page scripts beyond the standard `navigator.modelContext` surface.
- **`get_document(json)` can be large** — documents with embedded images carry
  `data:` URLs. Prefer `format:"text"` or `inspect_layout` scoping when you don't
  need the raw bytes. `inspect_layout` never emits image bytes (only `srcKind` +
  `srcLength`).
- **The document is the source of truth, not the agent.** In `examples/playground`
  the rebuild loop owns the document, so an agent's direct edits are replaced on the
  next code/data change — read & inspect tools are the natural fit there.
