# canvas-word

**A page-accurate, Word-class document editor for the browser, built on a
custom canvas rendering pipeline with [`@chenglou/pretext`](https://github.com/chenglou/pretext)
as the text layout engine.** The browser never lays out document text — we do,
deterministically, the way post-2021 Google Docs does.

**▶ Live demo: [doc-editor.forevka.dev](https://doc-editor.forevka.dev/)** — the
full feature showcase. Minimal offline embed:
[doc-editor.forevka.dev/examples/offline/](https://doc-editor.forevka.dev/examples/offline/).

---

## What this project wants to accomplish

### The thesis

Browsers are excellent at flowing text into elastic containers and terrible at
**paged documents**. A Word-class editor needs to answer questions the DOM
cannot answer honestly:

- *Where exactly does page 3 end?* — before rendering it, identically on every
  machine, so the scrollbar, the footer, and the printed sheet all agree.
- *Can this paragraph split here?* — widow/orphan rules, keep-with-next,
  row-level table breaks: decisions that require knowing every line's height
  before placing any of them.
- *What does this look like on paper?* — WYSIWYG that survives printing,
  because layout is computed in document coordinates, not viewport ones.

`contenteditable` editors (ProseMirror, Lexical, Slate) get selection, IME and
accessibility from the browser for free — and give up exactly the control
above. The alternative architecture is to **own the entire pipeline**: a
document model, a deterministic layout engine, a dumb paint layer, and an input
layer that rebuilds what `contenteditable` would have provided. That's the
architecture Google Docs migrated to (canvas-based rendering, 2021), and it is
the architecture here.

The historically prohibitive part — Unicode-correct, multi-script,
multi-style **line breaking** — is what pretext solves: pure TypeScript text
measurement and line layout without DOM reflows, with a per-line `maxWidth`
API that turned first-line indents into a one-liner and float text-wrap into a
placement mode rather than an engine rewrite.

### The goal, concretely

A typical Word user writes a real multi-page document — headings, lists,
tables, images, headers/footers, page numbers — **and doesn't notice it isn't
Word.** Same gestures (double-click into the header, Ctrl+Enter, Tab through
table cells, drag a column border), same break semantics (widows, orphans,
keep-with-next), same document model shape (runs + paragraph styles + sections,
deliberately aligned with OOXML so `.docx` maps mechanically in both directions
— both an import pipeline (`frontend/src/import/docx/`, see [IMPORT.md](./IMPORT.md))
and an export pipeline for `.docx` and PDF (`frontend/src/export/`, see
[EXPORT.md](./EXPORT.md)) have shipped).

### Non-goals

- Re-implementing the long tail of Word (VBA, SmartArt, equation editor,
  WordArt, charts). Each is its own product.
- Full OOXML fidelity. We target a **clean, honest subset** that round-trips;
  exotic constructs degrade with explicit warnings, never silently.
- Being a rich-text *component*. This is a paged document editor; if you need
  an inline rich-text field in a form, use Lexical.

### Where this sits next to the commercial editors

Teams that need an embeddable Word editor usually evaluate Syncfusion Document
Editor, OnlyOffice, and DevExpress Rich Text Editor. Those products cover more
of the Word long tail and ship enterprise support. They also charge a per-seat
license, and most expect a server (OnlyOffice Document Server, a Syncfusion
file-conversion endpoint, or a .NET backend) before the editor renders.

This project is the open-source, browser-native option. It ships under MIT,
paints to a `<canvas>` instead of leaning on `contenteditable`, carries zero
runtime dependencies, and runs fully offline; a backend is optional and only
buys live collaboration. It also goes **past** the incumbents in one place that
matters for generated reports: **nested content controls** (a w:sdt inside a
w:sdt, to any depth — including complex Word fields nested inside them)
round-trip faithfully, which the commercial editors flatten. The honest gaps
versus the incumbents are RTL/complex scripts, bundled CJK fonts, and
charts/equations (see Known limitations below).
A full side-by-side writeup lives at
[Best embeddable JS Word editors](https://forevka.dev/articles/best-embeddable-js-word-editors/).

---

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│  Input        mouse/keyboard · IME proxy · clipboard · keys  │
└───────────────┬─────────────────────────────────────────────┘
                │ commands (pure: state → transaction)
┌───────────────▼─────────────────────────────────────────────┐
│  Editor core  EditorState · transactions · undo (coalescing) │
└───────────────┬─────────────────────────────────────────────┘
                │ ops (each returns its exact INVERSE + pos-mapper)
┌───────────────▼─────────────────────────────────────────────┐
│  Model        Document → Block → Paragraph → Run             │
│               sections · stylesheet · tables · images        │
└───────────────┬─────────────────────────────────────────────┘
                │ revisions invalidate caches
┌───────────────▼─────────────────────────────────────────────┐
│  Layout       pretext line breaking → LineBoxes → pagination │
│               floats · row-chunked tables · margin bands     │
│               geometry queries (hit-test, caret, selection)  │
└───────────────┬──────────────────────┬───────────────────────┘
                │ absolute coordinates │
┌───────────────▼───────────────┐ ┌────▼──────────────────────┐
│  Paint (virtualized canvases) │ │  A11y mirror (hidden DOM) │
└───────────────────────────────┘ └───────────────────────────┘
```

One-way data flow: input → command → transaction → ops → new model →
incremental layout → paint + caret, coalesced per animation frame.

### The five invariants

Everything in this codebase follows from these. Violating one is a bug even if
nothing visibly breaks:

1. **Pretext is confined to `frontend/src/layout/`.** One adapter surface; API churn in
   a young dependency can never ripple.
2. **Paint never measures.** The paint layer receives absolutely-positioned
   fragments and calls `fillText` once per same-styled slice. If paint needs a
   width, layout failed to provide it.
3. **Every mutation is an invertible op.** Each `applyOp` returns the new
   document, the exact inverse op, and a position mapper. Undo is free,
   selections survive edits, and this seam is where track-changes and CRDT
   collaboration slot in later.
4. **Geometry is the single source of caret truth.** Hit-testing, caret rects,
   selection rects, line edges — all answered from the layout tree's indexed
   line list, never recomputed ad hoc.
5. **Caches are keyed by revision, invalidated by ops.** Two tiers:
   `prepareRichInline` per paragraph revision (the expensive measurement —
   width changes do NOT invalidate it), and `LineBox[]` per (revision, width).
   A keystroke re-breaks one paragraph; the rest is pagination arithmetic.

### Workspaces

The repo is an npm-workspace monorepo. The editor is the `frontend` workspace,
published as the `@forevka/wordcanvas` library; the document model and the OT
collaboration primitives live in `shared` so the Node `backend` can reuse them.

| Workspace | Owns |
|---|---|
| `shared/` | The document model (`shared/src/model/`: `document.ts` `ops.ts` `position.ts` `stylesheet.ts` `text.ts` `tableGrid.ts` `lists.ts`), serialization, and the OT change/transform/replay primitives (`change.ts` `transform.ts` `replay.ts` `ids.ts`) shared by editor and backend |
| `frontend/` | The editor itself (the `@forevka/wordcanvas` package) — layers below |
| `backend/` | Node HTTP/WS server: Postgres-backed change store, OT broadcast, server-side import/export, admin auth, integration tokens, session webhooks, OpenAPI/Swagger |
| `dashboard/` | Admin panel (Vite app): document upload → docId, session inspection |
| `examples/` | Standalone embeds consuming the published package (`embed-offline`, `embed-live`, `playground`, `embed-multi`) |
| `web/` | Caddy edge config for the deployed stack (see [DEPLOY.md](./DEPLOY.md)) |

### Editor layer map (`frontend/src/`)

| Directory | Owns | Key files |
|---|---|---|
| `layout/` | pretext integration, line caches, pagination + break rules, floats, tables, bands, footnotes, TOC, geometry queries, font metrics | `engine.ts` `geometry.ts` `prepareCache.ts` `layoutTree.ts` `metrics.ts` |
| `paint/` | Virtualized per-page canvases, DPR, selection/search rects, caret overlay, band dimming | `renderer.ts` |
| `child/` | Child documents: render/edit a content slice on canvas sharing the parent's live styles | `childDocument.ts` |
| `input/` | IME proxy, selection controller, object frames/resize, clipboard (HTML+plain), keymap | `imeProxy.ts` `selectionController.ts` `objectController.ts` `clipboard.ts` `keymap.ts` |
| `editor/` | Commands (pure), transactions, undo manager | `commands.ts` `state.ts` `undo.ts` |
| `a11y/` | Screen-reader mirror + live region | `mirror.ts` |
| `import/docx/` | .docx import pipeline (see [IMPORT.md](./IMPORT.md)) | `importDocx.ts` `pipeline.ts` … |
| `export/` | .docx + PDF export, DOM-free measure host (see [EXPORT.md](./EXPORT.md)) | `exportDocument.ts` `docx/` `pdf/` `shared/` |
| `builder/` | Fluent programmatic document composition + .docx templates (see [BUILDER.md](./BUILDER.md)) | `documentBuilder.ts` `template.ts` … |
| `sync/` | Live collaboration client: change recorder, OT sync client, presence | `SyncClient.ts` `changeRecorder.ts` `collab.ts` |
| `media/` | Image/media store (blob handling, backend upload bridge) | `store.ts` |
| `fonts/` | Metric-clone family map + baked metrics (see [FONTS.md](./FONTS.md)) | `clones.ts` |
| `ui/` | Ribbon menus, context menu, icons, table/SDT property editors | `contextMenu.ts` `tableProperties.ts` `sdtInspector.ts` |
| `app/` | App shell, runtime wiring, identity popup, busy overlay | `shell.ts` `runtime.ts` |
| `model/` | Sample + stress document generators (the model proper is in `shared/`) | `sampleDoc.ts` `stressDoc.ts` |
| `wordcanvas.ts` | The `WordCanvas` library entry (embeddable API) | |
| `editorApp.ts` / `main.ts` | Composition root + demo app chrome (toolbar, `?stress=`/`?docx=` loaders) | |

### Notable mechanisms

- **IME proxy** — a hidden, focusable `contenteditable` positioned at the
  caret. `beforeinput` is translated to commands and cancelled; CJK
  composition renders as *transient* underlined model edits outside the undo
  stack and commits as one insert on `compositionend`.
- **Story mode** — Word's double-click-into-header. The edited band renders
  *raw* (`{page}` literal, real block ids) so model offsets align; geometry
  queries scope to a per-band-per-page index; body dims; Esc restores the body
  selection. Ops are container-aware (`body | header | footer`).
- **Transient-op drag protocol** — image resize and column drags preview via
  ops with `origin:'transient'` (bypassing undo), then revert + commit ONE
  undoable op on mouseup. A whole drag = one Ctrl+Z.
- **Float text-wrap** — a `wrap:'square'` image registers a float rect instead
  of consuming height; following paragraphs re-break per line at the
  float-shrunk width for that line's y (with runCursor rollback when a line is
  rejected at a page boundary and re-broken at full width).
- **Whitespace-collapse offset maps** — pretext collapses runs of spaces; a
  tolerant fragment→model mapper emits a sparse offset map only when collapse
  happened, keeping caret↔pixel math exact (the "backspace deleted 'd' instead
  of 'o'" class of bug, fixed structurally).
- **Child documents** — `wc.createChild()` (or `editor.createChild()`) mints a
  lightweight sibling document that shares the parent's *live* style context
  (stylesheet, list/numbering defs, page section, content-control + field maps).
  It renders an arbitrary content slice — blocks, a fragment, runs, OOXML, or a
  named-style sample — through the **real** layout engine + canvas painter, so
  previews match the page exactly instead of an HTML approximation. The Home →
  Styles gallery swatches, the content-control inspector (read-only preview *and*
  the canvas-native editor, which commits edited blocks straight back through the
  `replaceSdt*` commands — no contentEditable round-trip), and the field
  constructor's result preview are all powered by it. Each child owns its own
  layout engine, so its cache can't collide with the parent's. A child also
  **proxies the parent's media** (its content-addressed "relationships"): it
  renders the parent's images, and a child editor can *add* new images
  (`childEditor.insertImage(bytes, mime)`) whose bytes register into the parent's
  shared store — so they persist, export, and survive commit-back. The
  content-control inspector exposes this as an "Insert Image…" affordance for
  block-level rich-text controls.

---

## Current feature matrix

Everything listed is implemented AND verified in-browser via scripted
Playwright checks (see `shots/` for visual evidence).

**Text & editing** — typing (native key events through the IME proxy), IME
composition with live preview, Enter/Backspace/Delete with Word's merge
semantics, soft selection (click/drag/double-word/triple-paragraph,
shift-extend, goal-X Up/Down, Home/End on *visual* lines, Ctrl+A/Home/End),
clipboard (HTML + plain, both directions, sanitized), cut, undo/redo with
typing coalescence, pending styles at a collapsed caret, page breaks
(Ctrl+Enter with Word's two-step Backspace removal).

**Lists** — bulleted and numbered, multilevel (9 levels; decimal/letter/roman
per level with OOXML %N marker patterns), counters reset on higher-level
increments, paint-only markers in hanging indents (line caches survive
renumbering), Tab/Shift+Tab promote/demote, Enter inherits / Enter-on-empty
exits, Backspace ladder (demote → leave list → page-break → merge), toolbar
toggles; `Document.lists` mirrors numbering.xml for the importer.

**Quick wins** — find & replace (Ctrl+F bar, match counter, prev/next,
replace + replace-all as ONE undo, live re-search on edits, orange highlight
channel); hyperlinks (paint-level blue+underline preserving the user's color,
hover pointer+tooltip, Ctrl+click opens, word-expansion at a collapsed caret,
HTML clipboard both ways); text highlight + sub/superscript (scaled font
measured honestly, baseline shift at paint); format painter (caret format →
next selection gesture; bare click applies paragraph format only — Word);
AutoCorrect (context-aware curly quotes, -- → em-dash, (c)/(r)/(tm) symbols,
riding the coalesced typing undo); soft line breaks (Shift+Enter inserts a
"\v" sentinel — paragraphs lay out as soft-break segments, each segment-final
line stays ragged under justify, empty trailing segments are caret-addressable
blank lines, Backspace deletes the break as one character, <br>/newline in the
clipboard both ways). **Everything planned in [ROADMAP.md](./ROADMAP.md) has
shipped** — including `.docx` and PDF export (see [EXPORT.md](./EXPORT.md)); only
raster browser-print was skipped in favor of the vector PDF path.

**Formatting** — bold/italic/underline/strikethrough, font family/size,
line spacing, letter-spacing, text color, alignment incl. **true
justification** (slack into `ctx.wordSpacing`; all 76 lines of the test
paragraph end at x=720.0 exactly), first-line indents, named styles
(gallery, apply/update-to-match-selection/create-from-selection, basedOn
chains, cascading modify — one undo reverts a cascade).

**Styles (Style Manager)** — a unified **Manage Styles** dialog (Home → Styles,
or right-click → Styles…) to create / edit / delete styles for every
OOXML-styleable entity, each with a **live canvas preview**: **paragraph** styles
(full char + para editor, cycle-safe based-on), first-class **character** styles
(`type` discriminator + per-run reference, gallery-applied, marked `ⓐ`), **list**
styles (per-level format/bullet/number-pattern/indent editor), and round-trippable
**table** styles with **conditional formatting** (header/total row, first/last
column, row/column banding) applied from right-click → Table Style. Every defined
style round-trips through `.docx` even before it's applied (matching Word); a
**"Show only styles in use"** gallery filter collapses large imported catalogs, and
a **Merge duplicates** action folds styles that are identical except for their name
into one (remapping references, one undo).

**Layout** — US-Letter pages with margins, line-level pagination
(paragraphs split mid-paragraph across pages), widow/orphan control (0
violations across 697 split paragraphs in the 1107-page audit),
keep-with-next, page breaks, exact scrollbar with canvas virtualization
(~3 live canvases for 1000+ pages).

**Tables** — editable cells (real Paragraphs; cells hold `Block[]` incl.
images scaled-to-fit and nested read-only tables), multi-paragraph cells,
Tab/Shift+Tab navigation (Tab in last cell appends a row), column resize by
border drag, row/col insert/delete, **cell merge (colSpan)** with span-aware
column ops, **row-level page breaks** (14-row table verified splitting 5+9).

**Images** — insert (body or inside cells), click-select with 8-handle frame,
proportional corner resize with live reflow, alignment, **square text-wrap**
(float left/right, text re-broken per line beside it), delete via key or
toolbar.

**Headers/footers** — rich `Block[]` stories (paragraphs/images/tables) laid
out per page with field tokens (`{page}` incl. roman/alpha formats, `{pages}`,
`{date}`, `{time}`), **story editing** via double-click with raw-token display,
scoped geometry, body dimming, Esc exit; **band tables are editable** (imported
footers are routinely a table beside a page-number paragraph); **first-page and
odd/even variants** ("Different first page" / "Different odd & even" in the 📐
panel) with per-page variant-aware story editing; **tall bands push the content
box** (a growing footer moves body text up live, like Word) and the band-edit
boundary/dim follow the real body edge.

**Sections & columns** — next-page section breaks (§⏎; OOXML sectPr-on-paragraph
model, "link to previous" band inheritance), per-section page setup via a
draggable **Page Layout** dialog (📐) with a live schematic preview and an
inches/cm toggle: Letter/A4/Legal **or custom** page size, orientation,
preset **or per-side** margins, header/footer band distances, page-number
restart, **page color** (`w:background`) and **page borders** (`w:pgBorders`),
all with per-page dimensions throughout paint/hit-testing. **Newspaper columns**
(per section; equal or **explicit per-column widths**, custom spacing, optional
**separator line**; flow fills columns then pages; floats clamp to their column;
Ctrl+Shift+Enter column break with the Word backspace ladder).

**Fields & TOC** — table of contents generated from Heading 1–3 paragraphs
(☰§ inserts at the caret or regenerates in place): entries are real paragraphs,
**page numbers are paint-only decorations resolved against the final page map**
(never stale, no fixpoint — entries reserve a right gutter so numbers can't
affect line breaking), dot leaders, Ctrl+click jumps to the heading.

**Footnotes** — a¹ inserts a superscript reference (its run text IS the note
number; inserting earlier renumbers everything in the same transaction) and
drops the caret into the new note; notes stack at the page bottom under a
⅓-width separator rule and are **editable in place** (they're real paragraphs
in the layout tree); a chunk referencing notes shrinks until text + notes
co-fit, so growing a note pushes body lines to the next page live.

**Break controls** — widow/orphan (always on), keep-with-next **chains**
(heading→subheading→body move as a group), keep-lines-together
(w:keepLines — paragraphs move whole instead of splitting), page and column
breaks.

**Content controls (SDT)** — inline w:sdt controls as first-class citizens:
rich/plain text, check box (click toggles ☐/☒), drop-down list & combo box
(click opens a chooser), date picker; gray placeholder text is selected whole
on entry and replaced by the first keystroke (Word); content/control locks
honored; Controls ribbon group inserts them, the importer maps `w:sdtPr`
losslessly (alias, tag, list items, date format, checkbox state, locks,
placeholder flag). **Nesting is first-class** — controls inside controls, and
block-level controls wrapping whole paragraphs/tables, to **any depth** —
carried as an ordered `sdtPath` ancestry that round-trips through `.docx`
import → edit → export without flattening or fragmenting. **Complex Word fields
nest inside controls** too (e.g. a DATE/PAGE/IF field as a control's value).
Hovering any control frames it (no caret move); nested controls draw concentric,
depth-graded frames with a breadcrumb tab (e.g. *Section › Appraisal Fee*).
**No other browser-native Word editor preserves nested content controls** — so a
generated report's section→field structure (the common C#/OOXML output) survives
a round-trip through the editor intact.

**Documents** — full `.docx` *import* (maps content controls, bookmarks and
hidden text; builds the style gallery from styles.xml `w:name` display names —
generated reports use opaque numeric styleIds; keeps **all** defined paragraph/
character/list/table styles so authored styles round-trip before use) and *export*
to `.docx` and PDF (see below), 1000-page stress generator (`?stress=N`).

**Export** — `.docx` (hand-rolled OOXML, the exact inverse of the importer, with
a `import(export(doc))` round-trip oracle that holds block count) and
**page-accurate PDF** (the editor's own layout engine replayed into pdfkit, so a
PDF page matches the canvas pixel-for-pixel). Both run isomorphically — the same
DOM-free pipeline in a browser Web Worker and on the Node backend, over bundled
metric-clone fonts so editor, browser export and server export paginate
identically. Embedders can **save to their own pipeline**: `exportDocx()` /
`exportPdf()` on the instance return a `Blob` for a custom Save button, and an
`onSave` constructor option re-routes the toolbar's Export buttons to a callback
(`Blob` + bytes + format + warnings) instead of a browser download. See
[EXPORT.md](./EXPORT.md) and [FONTS.md](./FONTS.md).

**Programmatic generation** — a fluent document-builder API
(`@forevka/wordcanvas/builder`): compose documents in JS/TS instead of a
C#/OOXML backend — `DocumentBuilder.fromTemplate(docx)` carries a template's
styles/page setup/bands, chain `.paragraph(data.title).withStyle("Heading1")`,
tables/images/lists/headers/footers, feed a JSON data model, and live-preview
via `WordCanvas.setDocument(doc)` (declarative rebuild on data change,
zoom/scroll preserved). Runs in the browser and in Node. Interactive
playground at `examples/playground`. See [BUILDER.md](./BUILDER.md).

**Collaboration** — operational-transform sync over a WebSocket backend
(`shared/` holds the change/transform/replay primitives; the editor's `sync/`
records edits and the `backend/` broadcasts them against a Postgres change
store). Opening a doc with a `backendUrl` set auto-publishes it (gzipped
snapshot + parallel media upload) and returns a share link; remote
collaborators' carets **and selections** render with attribution; an activity
panel shows who created/edited and when; identity is owned by the embedder.

**Track changes & comments (review)** — a three-way mode switch (Editing /
Suggesting / Viewing) in the ribbon header. In **Suggesting** mode every edit
becomes a tracked, attributed proposal: insertions paint as author-colored
underlines, deletions as strikethroughs (**non-destructive — the text stays in
the model until accepted**), with margin change-bars. **Comments** thread on a
text range with reply + resolve, **@-mentions** (type `@` to tag from an
embedder-supplied `knownUsers` roster; resolved identities are recorded per
comment and the backend fires a `comment.mention` webhook to notify third-party
systems), a Google-Docs-style floating composer bubble
(an "add a comment" chip auto-appears on selection in suggest mode), and a docked
**Review pane** (right side, like the outline drawer) that lists suggestions and
comment threads with Accept / Reject (and accept-/reject-all), click-to-scroll
navigation, and live counts. Everything syncs in real time on a dedicated review
channel (idempotent, causally delivered) and **rehydrates on join** — a user
opening a shared document sees existing feedback, anchored correctly even after
later edits. Crucially this is an **overlay extension**: the OOXML-faithful core
model and the `.docx`/PDF writers never see it — anchors ride the same
`mapPosition` rebasing as bookmarks, and the default export *bakes* the overlay
into a clean document (reject-all = the original baseline). See [REVIEW.md](./REVIEW.md).

**Agent tools (WebMCP)** — opt-in (`agentTools: true`), the editor exposes itself to
AI agents over [WebMCP](https://webmcp.dev) (the standard `navigator.modelContext`
API). An agent connected via the WebMCP browser extension / Chrome DevTools MCP gets
tools to **read & inspect** (`get_document`, `search_document`, `inspect_layout` — a
page/line/fragment **geometry dump** for debugging "renders weirdly" reports —
`get_document_stats`), **suggest & comment** (mode switch, threaded comments,
tracked-change accept/reject), and **directly edit** (find-anchored
replace/insert/format, alignment, undo/redo). Capability-gated and namespaced for
multiple editors per page; the WebMCP polyfill is lazy-loaded only when enabled, so
the zero-runtime-dependency package is unchanged for everyone else. See
[WEBMCP.md](./WEBMCP.md).

**Develop mode** — opt-in (`develop: true`), the ribbon gains a dedicated
**Developer** tab whose *Inspect document tree* button opens a floating,
devtools-Elements-style panel over the parsed `Document` model. Browse the tree
(body blocks → runs, tables → rows → cells, header/footer bands, footnotes, and the
side-tables: styles, lists, table styles, content controls, fields, bookmarks);
**hover a node** to paint a highlight box over its region on the canvas; **hover the
page** to reveal the matching node in the tree (reverse sync); **click** to select
and scroll to it; and **select** any node to read its properties + raw JSON. It's a
pure debugging aid — gated twice over: the tab only exists when the flag is set, and
nothing runs until the developer opens the panel from it. Leave it off for
production embeds.

**Bookmarks** — character-range bookmarks (body, table cells, headers/footers)
that rebase live as the document is edited; a Bookmarks panel lists them with
Go-To navigation and management ops; they back TOC and cross-reference targets.

**Hidden text** — `w:vanish` runs are preserved through round-trips but never
laid out or painted (matches Word's hidden bookmark-anchor paragraphs).

**Mobile / touch** — on-screen-keyboard input through the IME proxy, pointer
(touch) resize handles for objects, and a responsive ribbon that collapses on
narrow viewports.

## Verified performance

| Metric | Result |
|---|---|
| Cold layout, 4950 blocks / 1107 pages | ~1.15 s (≈0.23 ms/paragraph, all `prepareRichInline`) |
| Warm full relayout (line cache) | **~3 ms** |
| Keystroke at page ~550 of 1107 (full pipeline) | min 4.3 / **median 5.2** / max 14.7 ms |
| Scroll-jump to page 550 | settles < 10 ms, 4 live canvases |
| Break-rule audit over 1107 pages | 0 orphan / 0 widow / 0 keep-with-next violations |

## Getting started

This is an npm-workspace monorepo; run everything from the repo root.

```sh
npm install
npm run dev          # editor dev server (frontend), open the printed localhost URL
npm run typecheck    # typechecks shared + frontend + backend + dashboard
npm run test         # vitest across shared + frontend + backend
```

Other workspaces and the full stack:

```sh
npm run build:lib     # build the @forevka/wordcanvas library bundle (dist-lib)
npm run dev:online    # one command: Postgres + backend + frontend in ONLINE mode
                      #   (Share button, live collaboration, track-changes/comments sync)
npm run db:up         # start Postgres + apply migrations (docker compose; backend runs separately)
npm run dev:dashboard # admin dashboard dev server
npm run dev:example   # the embed-live example (consumes the built library)
npm run dev:playground # document-builder playground (code + JSON → live preview)
npm run dev:multi     # multiple editors on one page (consumes the built library)
```

- `/?stress=1000` — generate a ~1000-page perf-probe document and log timings.
- `/?docx=<url>` — fetch and import a .docx.
- `window.__cw` — dev hook exposing `{ doc, tree, engine, editor, … }` for
  in-browser verification scripts.
- Toolbar: styles gallery (✎ update-to-match, ✚ Manage Styles, ⏷ show-only-in-use
  filter), Style Manager dialog (paragraph/character/list/table styles), font,
  size, spacing, undo/redo, B/I/U/S, alignment, insert image/table, table
  row/col ops, merge/unmerge, image wrap modes, open .docx.
- Keys: Ctrl+B/I/U, Ctrl+Z/Y, Ctrl+Enter (page break), Tab in tables,
  Esc (exit header mode / deselect object), Ctrl+A/Home/End, Shift+arrows.

## Roadmap

See **[ROADMAP.md](./ROADMAP.md)** for the full feature plan and what shipped.
Everything in it is done; the editor covers ~95% of everyday Word usage.

## Known limitations (honest list)

- RTL/bidi caret affinity unhandled (pretext renders bidi; editing it is open).
- Real-IME tested via synthetic events only — needs a hands-on CJK keyboard pass.
- A11y mirror is minimal (aria-label sync + live region; full mirror planned).
- No repeat-header-row on table chunks; a vertical-merge (rowSpan) table taller
  than a page splits at a row boundary, leaving the merged cell on the first page
  only (a table that fits whole is moved intact to the next page instead).
- Only Latin metric-clone fonts are bundled, so CJK/complex scripts render as
  tofu in PDF export (same gap as the importer).
- Raster browser-print was skipped — use PDF export instead.
- `font-feature-settings` / optical sizing unsupported (pretext limitation).
- Child-document previews render a single content flow (no pagination, cropped to
  the surface); live page fields (`{page}`/`{numpages}`) show sample values since a
  child has no pagination context, and a re-rendered preview can briefly blank for a
  frame while its page canvas remounts. Images resolve through the parent's shared
  media store (existing and child-added), but the `kind:"ooxml"` slice path itself
  is still media-free (a raw fragment has no relationships, so images by `r:id` in
  that string won't resolve — pass `kind:"blocks"` for media-bearing content).
- Track-changes covers insert/delete/format, structural edits (paragraph
  split/merge, block & table row/column ops), and paste/cross-paragraph deletes
  (decomposed per-op, bundled as one accept/reject unit). Remaining gaps: comment
  bodies are plain text (rich is still V2), overlapping suggestions resolve by a
  documented precedence rather than partial-overlap handling, and a structural
  suggestion's stored inverse isn't re-indexed under heavy intervening edits
  before it's resolved.
- Table styles bake their effective per-cell formatting onto cells when **applied
  in the editor** (so header shading/banding/borders show and export correctly);
  on **import**, the style definition and reference round-trip but conditional
  bands aren't re-baked onto cells, and a style's character/paragraph conditional
  formatting round-trips without being applied to cell text. Removing the import
  "keep all defined styles" pruning means a heavy third-party template surfaces its
  full style catalog in the gallery (use the "show only styles in use" filter).

## Implementation history

| Milestone | Delivered |
|---|---|
| 1 | Read-only renderer: model → pretext → paginated canvas, virtualization |
| 2 | Break rules (widow/orphan/keep-with-next), 1107-page stress audit |
| 3 | Geometry: hit-testing, caret, selection, keyboard nav, lazy cluster advances |
| 4 | Editing: invertible ops, undo coalescence, IME proxy, a11y, line caches |
| 5 | Rich text: justification, pending styles, HTML clipboard, toolbar |
| 6 | Word furniture: images, tables, rich headers/footers, story editing |
| 7 | Styles system, object manipulation (resize/wrap), cell merge, row-level table breaks, floats, `Block[]` cells, page breaks |
| 8 | Roadmap close-out: lists, find/replace, hyperlinks, sub/super, format painter, autocorrect, soft breaks, sections, columns, header variants, fields/TOC, footnotes |
| 9 | `.docx` + page-accurate PDF export (isomorphic browser/Node, bundled metric-clone fonts) |
| 10 | Product layer: OT collaboration + Postgres backend, admin dashboard, integration tokens + session webhooks, `@forevka/wordcanvas` npm package, containerized VPS deploy behind Caddy |
| 11 | Track changes & comments: edit/suggest/view modes, non-destructive tracked insert/delete/format, threaded comments with a floating composer, docked Review pane (accept/reject + scroll-to), real-time review sync with join rehydration, default export bake — all an overlay extension that never touches the OOXML core (see [REVIEW.md](./REVIEW.md)) |
