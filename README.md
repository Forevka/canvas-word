# canvas-word

A page-accurate, Word-class document editor for the browser. It paints to
`<canvas>` and owns the entire pipeline — document model, deterministic layout,
dumb paint layer, input layer — instead of leaning on `contenteditable`. Text
layout runs on [`@chenglou/pretext`](https://github.com/chenglou/pretext), the
same architecture Google Docs moved to in 2021.

**▶ Live demo: [doc-editor.forevka.dev](https://doc-editor.forevka.dev/)** ·
minimal embed: [/examples/offline/](https://doc-editor.forevka.dev/examples/offline/)

## Related projects

- [**dom-docx**](https://github.com/floodtide/dom-docx) — converts semantic HTML
  fragments into native, editable Word documents (OOXML). Complementary to
  canvas-word: generate a `.docx` from HTML with dom-docx, then open and edit it
  here.

## Why canvas instead of contenteditable

Browsers flow text into elastic containers well and handle paged documents
badly. A Word-class editor has to know where page 3 ends before rendering it,
whether a paragraph can split at a given line (widows, orphans,
keep-with-next), and what the page looks like on paper — all in document
coordinates, identically on every machine. `contenteditable` editors
(ProseMirror, Lexical, Slate) get selection, IME, and accessibility for free
but can't answer those questions. Owning the pipeline is the trade this project
makes; pretext supplies the hard part, Unicode-correct multi-script line
breaking, without DOM reflows.

**Goal:** a Word user writes a real multi-page document — headings, lists,
tables, images, headers/footers, page numbers — with the same gestures, break
semantics, and OOXML-aligned model, and doesn't notice it isn't Word.

**Non-goals:** the Word long tail (VBA, SmartArt, charts); full OOXML fidelity
(we target a clean subset that round-trips, degrading exotic constructs with
explicit warnings); being an inline rich-text component (use Lexical for that).

**Versus the commercial editors** (Syncfusion, OnlyOffice, DevExpress): this is
the MIT, browser-native, zero-runtime-dependency, offline option — no server
required (a backend only adds live collaboration). It also preserves **nested
content controls** (`w:sdt` inside `w:sdt`, to any depth) that the commercial
editors flatten, plus RTL/bidi/CJK editing and MathML equations. Full writeup:
[Best embeddable JS Word editors](https://forevka.dev/articles/best-embeddable-js-word-editors/).

## Architecture

```
Input  → command → transaction → ops → new model → incremental layout → paint + caret
(mouse/kbd,        (pure:         (each returns    (Document →           (pretext line-break
 IME, clipboard)    state→txn)     its inverse)     Block → Para → Run)   → LineBoxes → pages)
```

One-way data flow, coalesced per animation frame. A hidden-DOM a11y mirror runs
alongside paint.

### The five invariants

Violating one is a bug even if nothing visibly breaks:

1. **Pretext is confined to `frontend/src/layout/`** — one adapter surface, so
   API churn in a young dependency can't ripple.
2. **Paint never measures.** It receives absolutely-positioned fragments and
   calls `fillText` per same-styled slice. Needing a width means layout failed.
3. **Every mutation is an invertible op.** `applyOp` returns the new document,
   the exact inverse, and a position mapper — undo is free, selections survive
   edits, and this is where track-changes and CRDT collaboration slot in.
4. **Geometry is the single source of caret truth.** Hit-testing, caret rects,
   selection, line edges all come from the layout tree's indexed line list.
5. **Caches are keyed by revision, invalidated by ops.** `prepareRichInline`
   per paragraph revision (width changes don't invalidate it) and `LineBox[]`
   per (revision, width). A keystroke re-breaks one paragraph; the rest is
   pagination arithmetic.

### Workspaces

npm-workspace monorepo. The editor is `frontend` (published as
`@forevka/wordcanvas`); the model and OT primitives live in `shared` so the
Node `backend` can reuse them.

| Workspace | Owns |
|---|---|
| `shared/` | Document model (`document.ts` `ops.ts` `position.ts` `stylesheet.ts` `tableGrid.ts` `lists.ts`), serialization, OT primitives (`change.ts` `transform.ts` `replay.ts`) |
| `frontend/` | The editor — the `@forevka/wordcanvas` package |
| `backend/` | Node HTTP/WS server: Postgres change store, OT broadcast, server-side import/export, admin auth, session webhooks, OpenAPI |
| `dashboard/` | Admin panel (Vite): document upload → docId, session inspection |
| `examples/` | Standalone embeds consuming the published package |
| `web/` | Caddy edge config for the deployed stack (see [DEPLOY.md](./DEPLOY.md)) |

### Editor layer map (`frontend/src/`)

| Directory | Owns | Key files |
|---|---|---|
| `layout/` | pretext integration, line caches, pagination + break rules, floats, tables, footnotes, TOC, geometry, font metrics | `engine.ts` `geometry.ts` `prepareCache.ts` `layoutTree.ts` `metrics.ts` |
| `paint/` | Virtualized per-page canvases, DPR, selection/search rects, caret overlay | `renderer.ts` |
| `child/` | Child documents: render/edit a content slice sharing the parent's live styles | `childDocument.ts` |
| `input/` | IME proxy, selection controller, object frames/resize, clipboard, keymap | `imeProxy.ts` `selectionController.ts` `objectController.ts` `clipboard.ts` |
| `editor/` | Commands (pure), transactions, undo manager | `commands.ts` `state.ts` `undo.ts` |
| `a11y/` | Screen-reader mirror + live region | `mirror.ts` |
| `import/docx/` | .docx import (see [IMPORT.md](./IMPORT.md)) | `importDocx.ts` `pipeline.ts` |
| `export/` | .docx + PDF export, DOM-free measure host (see [EXPORT.md](./EXPORT.md)) | `exportDocument.ts` `docx/` `pdf/` |
| `builder/` | Fluent document composition + .docx templates (see [BUILDER.md](./BUILDER.md)) | `documentBuilder.ts` `template.ts` |
| `codegen/` | Reverse builder: `.docx` → DocumentBuilder code (see [CODEGEN.md](./CODEGEN.md)) | `emit.ts` `index.ts` |
| `sync/` | Collaboration client: change recorder, OT sync, presence | `SyncClient.ts` `changeRecorder.ts` |
| `media/` | Image/media store (blobs, backend upload bridge) | `store.ts` |
| `fonts/` | Metric-clone family map + baked metrics (see [FONTS.md](./FONTS.md)) | `clones.ts` |
| `ui/` | Ribbon, context menu, icons, table/SDT property editors | `contextMenu.ts` `tableProperties.ts` |
| `app/` | App shell, runtime wiring | `shell.ts` `runtime.ts` |
| `model/` | Sample + stress document generators | `sampleDoc.ts` `stressDoc.ts` |

### Notable mechanisms

- **IME proxy** — a hidden `contenteditable` at the caret. `beforeinput`
  becomes commands and is cancelled; CJK composition renders as transient
  underlined edits outside the undo stack, committing as one insert on
  `compositionend`.
- **Story mode** — double-click into header/footer. The edited band renders raw
  (`{page}` literal, real block ids) so offsets align; the body dims; Esc
  restores selection. Ops are container-aware (`body | header | footer`).
- **Transient-op drag** — image resize and column drags preview via
  `origin:'transient'` ops (bypassing undo), then revert and commit one undoable
  op on mouseup. A whole drag is one Ctrl+Z.
- **Float text-wrap** — a `wrap:'square'` image registers a float rect instead
  of consuming height; following paragraphs re-break per line at the shrunk
  width, with rollback when a line is rejected at a page boundary.
- **Whitespace-collapse offset maps** — pretext collapses space runs; a sparse
  fragment→model offset map (emitted only on collapse) keeps caret↔pixel math
  exact.
- **Child documents** — `wc.createChild()` mints a lightweight sibling that
  shares the parent's live style context and renders an arbitrary slice through
  the real layout engine + painter, so previews match the page. It powers the
  Styles gallery, the content-control inspector's canvas-native editor, and the
  field constructor preview. A child proxies the parent's media and can add
  images that register into the shared store.

## Feature matrix

Everything below is implemented and verified in-browser via scripted Playwright
checks (`shots/` holds visual evidence).

**Text & editing** — typing through the IME proxy, IME composition with live
preview, Enter/Backspace/Delete with Word's merge semantics, full soft
selection (drag/double/triple, shift-extend, goal-X, visual-line Home/End,
Ctrl+A/Home/End), sanitized HTML+plain clipboard both ways, cut, undo/redo with
typing coalescence, pending styles at a collapsed caret, Ctrl+Enter page breaks.

**Formatting** — bold/italic/underline/strike, font family/size, line and
letter spacing, color, alignment including true justification (slack into
`ctx.wordSpacing`), first-line indents. **Contextual floating toolbars** pop above
the caret/selection — one priority-based manager shows the most relevant of a
built-in **text format** bar (Word's selection toolbar: font, size, B/I/U/S, colour,
highlight — configurable via `floatingToolbar`; see
[`examples/floating-toolbar`](./examples/floating-toolbar)) plus bars for a selected
**image** / **equation**, a multi-cell **table** selection, a **hyperlink**, a
**comment/suggestion**, a **footnote/endnote** marker, the **TOC**, a **list item**, and
an empty-paragraph ＋ **insert menu**. Embedders can register their own via the
`contextToolbars` option (see
[`examples/context-toolbars`](./examples/context-toolbars)).

**Styles (Style Manager)** — one **Manage Styles** dialog to create/edit/delete
styles for every OOXML-styleable entity, each with a live canvas preview:
paragraph (cycle-safe based-on), character (per-run, marked `ⓐ`), list
(per-level format), and table styles with conditional formatting (header/total
row, first/last column, banding). Every defined style round-trips through
`.docx` before it's applied; a "show only styles in use" filter and a "merge
duplicates" action manage large imported catalogs.

**Lists** — bulleted and numbered, 9 multilevel levels (decimal/letter/roman,
OOXML `%N` patterns), counter resets, paint-only markers in hanging indents,
Tab/Shift+Tab promote/demote, Enter inheritance and Enter-on-empty exit,
Backspace ladder. `Document.lists` mirrors numbering.xml.

**Layout** — US-Letter pages with margins, line-level pagination (mid-paragraph
splits), widow/orphan control, keep-with-next, exact scrollbar with canvas
virtualization (~3 live canvases for 1000+ pages).

**Tables** — editable `Block[]` cells (images, nested read-only tables),
multi-paragraph cells, Tab navigation (Tab in last cell appends a row), column
resize by border drag, row/col insert/delete, cell merge (colSpan) with
span-aware ops, row-level page breaks, and content-driven autofit (AutoFit to
Contents/Window; `w:tcW`/`w:tblLayout`/`w:tblW` round-trip).

**Images** — insert in body or cells, 8-handle select frame, proportional
resize with live reflow, alignment, square text-wrap (float left/right).

**Headers/footers** — rich `Block[]` stories laid out per page with field
tokens (`{page}` with roman/alpha, `{pages}`, `{date}`, `{time}`), story
editing via double-click, editable band tables, first-page and odd/even
variants, tall bands that push the content box live.

**Sections & columns** — next-page section breaks (OOXML sectPr-on-paragraph),
per-section page setup via a draggable **Page Layout** dialog (Letter/A4/Legal
or custom size, orientation, margins, band distances, page-number restart, page
color `w:background`, page borders `w:pgBorders`), and newspaper columns (equal
or per-column widths, custom spacing, optional separator, column breaks).

**Fields & TOC** — table of contents from Heading 1–3; entries are real
paragraphs, page numbers are paint-only decorations resolved against the final
page map (never stale), with dot leaders and Ctrl+click jump.

**Footnotes** — superscript reference (its run text is the note number;
inserting earlier renumbers in the same transaction), notes stack at page
bottom under a separator and are editable in place; growing a note pushes body
lines to the next page live.

**Content controls (SDT)** — inline `w:sdt` as first-class citizens: rich/plain
text, check box, drop-down/combo, date picker; placeholder text selected whole
on entry; locks honored; the importer maps `w:sdtPr` losslessly. **Nesting is
first-class** — controls inside controls and block-level controls wrapping whole
paragraphs/tables, to any depth, carried as an `sdtPath` ancestry that
round-trips through import → edit → export without flattening. Complex Word
fields nest inside controls too. **No other browser-native Word editor preserves
nested content controls**, so a generated report's section→field structure
survives a round-trip intact.

**Documents** — full `.docx` import (content controls, bookmarks, hidden text;
builds the style gallery from `w:name`, keeps all defined styles) and export to
`.docx` and PDF; 1000-page stress generator (`?stress=N`).

**Export** — `.docx` (hand-rolled OOXML, exact inverse of the importer, with an
`import(export(doc))` round-trip oracle) and **page-accurate PDF** (the layout
engine replayed into pdfkit, pixel-for-pixel with the canvas). Both run
isomorphically in a browser Web Worker and on the Node backend over bundled
metric-clone fonts. Embedders can save to their own pipeline: `exportDocx()` /
`exportPdf()` return a `Blob`, and an `onSave` option re-routes the toolbar
Export buttons. See [EXPORT.md](./EXPORT.md), [FONTS.md](./FONTS.md).

**Programmatic generation** — a fluent builder (`@forevka/wordcanvas/builder`):
compose documents in JS/TS instead of a C#/OOXML backend.
`DocumentBuilder.fromTemplate(docx)` carries a template's styles/page
setup/bands; feed a JSON data model and live-preview via
`WordCanvas.setDocument(doc)`. Runs in browser and Node. Playground at
`examples/playground`. See [BUILDER.md](./BUILDER.md).

**Reverse builder** — the inverse (`@forevka/wordcanvas/codegen`):
`docxToBuilderCode(bytes)` emits editable TypeScript that calls the fluent
builder to reconstruct a document, turning a visually-authored `.docx` into a
code template. Fields the builder can't yet express are reported, not dropped.
See [CODEGEN.md](./CODEGEN.md).

**Collaboration** — operational-transform sync over a WebSocket backend.
Opening a doc with `backendUrl` set auto-publishes it (gzipped snapshot +
parallel media upload) and returns a share link; remote carets and selections
render with attribution; an activity panel shows who edited when; identity is
owned by the embedder.

**Track changes & comments** — a three-way mode switch (Editing / Suggesting /
Viewing). In Suggesting mode edits become attributed proposals (author-colored
underlines for inserts, strikethroughs for deletes, **non-destructive** until
accepted, with margin change-bars). Comments thread on a range with
reply/resolve and @-mentions (from an embedder `knownUsers` roster; the backend
fires a `comment.mention` webhook), a floating composer, and a docked Review
pane with Accept/Reject (+ all) and click-to-scroll. Everything syncs in real
time and rehydrates on join. It's an **overlay extension** — the OOXML core and
the writers never see it; default export bakes the overlay into a clean document
(reject-all = the original baseline). See [REVIEW.md](./REVIEW.md).

**Agent tools (WebMCP)** — opt-in (`agentTools: true`), the editor exposes
itself to AI agents over [WebMCP](https://webmcp.dev)
(`navigator.modelContext`): read/inspect (`get_document`, `search_document`,
`inspect_layout`), suggest/comment, and directly edit (find-anchored
replace/insert/format, undo/redo). Capability-gated, namespaced, lazy-loaded
only when enabled. See [WEBMCP.md](./WEBMCP.md).

**Develop mode** — opt-in (`develop: true`), a Developer ribbon tab opens a
devtools-Elements-style panel over the parsed `Document`: browse the tree, hover
a node to highlight its canvas region, hover the page for reverse sync, click to
select and read properties + raw JSON. Gated twice — the tab only exists when
the flag is set. Leave it off for production.

**Also** — bookmarks (character-range, rebase live, back TOC and
cross-references); hidden text (`w:vanish` preserved but never laid out); math
equations (MathML-native, STIX Two Math typesetting, OMML round-trip, LaTeX
input); mobile/touch (on-screen keyboard via IME proxy, touch resize handles,
responsive ribbon).

## Verified performance

| Metric | Result |
|---|---|
| Cold layout, 4950 blocks / 1107 pages | ~1.15 s (≈0.23 ms/paragraph) |
| Warm full relayout (line cache) | **~3 ms** |
| Keystroke at page ~550 of 1107 | median **5.2 ms** (min 4.3 / max 14.7) |
| Scroll-jump to page 550 | settles < 10 ms, 4 live canvases |
| Break-rule audit over 1107 pages | 0 orphan / 0 widow / 0 keep-with-next violations |

## Getting started

npm-workspace monorepo; run everything from the repo root.

```sh
npm install
npm run dev          # editor dev server (frontend)
npm run typecheck    # shared + frontend + backend + dashboard
npm run test         # vitest across shared + frontend + backend
```

Full stack and other workspaces:

```sh
npm run build:lib      # build the @forevka/wordcanvas bundle (dist-lib)
npm run dev:online     # Postgres + backend + frontend in ONLINE mode (Share, collab, review sync)
npm run db:up          # start Postgres + apply migrations (docker compose)
npm run dev:dashboard  # admin dashboard
npm run dev:example    # embed-live example (consumes the built library)
npm run dev:playground # document-builder playground
npm run dev:multi      # multiple editors on one page
```

- `/?stress=1000` — generate a ~1000-page perf-probe document and log timings.
- `/?docx=<url>` — fetch and import a .docx.
- `window.__cw` — dev hook exposing `{ doc, tree, engine, editor, … }`.
- Keys: Ctrl+B/I/U, Ctrl+Z/Y, Ctrl+Enter (page break), Tab in tables, Esc
  (exit header mode / deselect), Ctrl+A/Home/End, Shift+arrows.

## Roadmap & status

Everything in [ROADMAP.md](./ROADMAP.md) has shipped; the editor covers ~95% of
everyday Word usage. Only raster browser-print was skipped, in favor of the
vector PDF path.

## Known limitations

- Real IME tested via synthetic events only — needs a hands-on CJK keyboard pass.
- A11y mirror is minimal (aria-label sync + live region).
- Header rows repeat on continuation pages (`w:tblHeader`), except in a table
  with a vertical merge. A rowSpan table taller than a page splits at a row
  boundary, leaving the merged cell on the first page (one that fits whole moves
  intact to the next page instead).
- **Autofit table export is hint-based.** The writer emits
  `w:tblLayout="autofit"` and lets Word re-autofit from the `w:gridCol`
  snapshot rather than writing solved widths; a non-Word consumer may not
  reproduce the editor's columns (our PDF export renders through the layout
  engine, so it's exact). Import adopts autofit only when explicitly declared.
- Bundled font fallbacks default on for PDF export (zero config): CJK (Noto Sans
  SC), Arabic (Noto Sans Arabic, with shaping), and Hebrew (Noto Sans Hebrew),
  so complex-script text no longer exports as tofu. The CJK face is a
  common-Chinese subset — rarer hanzi or fuller CJK coverage still need a
  registered font (`cjk.fallbackFont` / `fonts`); each fallback is a single
  Regular face (no bold/italic outlines). DOCX export embeds no fonts: it keeps
  the document's own family names and relies on the reader (Word) having them.
- **Math equations are a focused subset** — Presentation MathML, STIX Two Math
  typesetting, OMML round-trip (inline + display), MathML/LaTeX input (common
  command set, no AsciiMath). Stretchy delimiters use glyph size-scaling, not
  the MATH table's piecewise assembly; unsupported OMML is surfaced, not dropped.
- Child-document previews render a single content flow (no pagination); page
  fields show sample values; the `kind:"ooxml"` slice path is media-free (pass
  `kind:"blocks"` for images by `r:id`).
- Track changes covers insert/delete/format and structural edits (split/merge,
  block & table row/col ops, cross-paragraph paste). Gaps: comment bodies are
  plain text, overlapping suggestions resolve by documented precedence, a
  structural suggestion's inverse isn't re-indexed under heavy intervening edits.
- Table styles bake per-cell formatting when applied in the editor; on import
  the definition and reference round-trip but conditional bands aren't re-baked.
