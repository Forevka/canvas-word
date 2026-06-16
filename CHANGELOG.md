# Changelog

All notable changes to **`@forevka/wordcanvas`** are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.7.1] — 2026-06-16

### Fixed
- **Opening a second document no longer merges it with the previous one.** The layout
  engine is shared across documents and caches laid-out lines by block id; because the
  docx importer re-mints block ids from `i0` on every import, the freshly opened
  document collided with the prior one's cache and rendered the two merged. Replacing
  the document now clears the engine's caches first, so each opened document renders on
  its own.
- **Resizing a table column no longer scrolls the view back to the caret.** When the
  caret sat in a paragraph elsewhere (e.g. on another page), dragging a column grip
  yanked the viewport back to the caret on every drag tick. The editor now only
  scrolls the caret into view when a mutation actually moves it, so selection-preserving
  edits like a column resize leave the scroll position untouched.

## [0.7.0] — 2026-06-16

### Added
- **Copy a rectangular table-cell selection.** Selecting whole cells and copying now
  writes an HTML `<table>` (preserving `colspan`/`rowspan`) alongside tab-separated
  plain text, so it pastes faithfully into Word / Google Docs and into spreadsheets.
- **First-load progress callback.** New optional `onLoadProgress` option (and exported
  `LoadProgress` type) reports cold-load progress so embedders can show a loader while
  the big chunks stream: the editor JS chunk download (`phase: "bundle"`) and the
  bundled ~9 MB font fetch (`phase: "fonts"`, a smooth size-weighted bar), finishing
  with `phase: "ready"`. `percent` is an overall 0..1 monotonic value to drive a
  progress bar. Wired into the bundled examples and the root demo page.

### Fixed
- **Table cell text no longer overflows its column.** Cell paragraphs wrap at the
  cell's inner width minus their own left/right indent (matching body wrapping), and
  cell content is clipped to the inner box in both the canvas and PDF painters — so
  indented text wraps in-cell instead of drawing over the right border or into the
  neighbouring column.
- **Cell selection stays visible over shaded cells.** Cell background fills are now
  painted beneath the selection/search highlights, and the selection uses an adaptive
  blend, so it no longer disappears on a coloured (e.g. blue) table header.
- **Whole-cell copy works for column-spanning cells.** Dragging across a merged cell
  no longer trips a false cross-cell selection that left a partial range, so Ctrl+C
  copies the full cell content.

## [0.6.1] — 2026-06-15

### Changed
- Internal maintainability refactor with no public API or behavior change:
  consolidated the duplicated OOXML value maps and run/paragraph serialization,
  extracted the SDT chooser and comment composer into dedicated controllers, added
  a shared `injectCssOnce` helper, and switched modal Escape handling to
  `AbortController`-based listener cleanup.

### Fixed
- **Layout caches** now evict entries for deleted blocks instead of growing
  unbounded across a long editing session.
- **Export**: non-finite style values can no longer serialize to an invalid
  `w:val="NaN"`; the cell `w:vMerge` restart is emitted in-place rather than via
  fragile string patching.
- **Suggestions / agent input**: the track-changes interceptor no longer anchors a
  suggestion to a missing block, and the WebMCP `select_range` tool validates block
  ids against the live document and clamps offsets.

## [0.6.0] — 2026-06-15

### Added
- **AI agent tools over WebMCP.** New opt-in `agentTools` option exposes the editor
  to AI agents through the standard `navigator.modelContext` API (polyfilled by
  `@mcp-b/global`, lazy-loaded so non-agent embedders pay nothing — the package
  stays zero-runtime-dependency). Tools span three capability buckets: **read &
  inspect** (`get_document`, `get_selection`, `search_document`, `inspect_layout` —
  a page/line/fragment geometry dump for debugging rendering issues —
  `get_document_stats`), **suggest & comment**, and **direct edits**
  (find-anchored replace/insert/format, alignment, `select_range`, undo/redo,
  `set_document`). Capability-gated and namespaceable for multiple editors per page.
  Enabled in all bundled examples; fully documented in
  [WEBMCP.md](./WEBMCP.md).
- **Generic custom OOXML fields** with host-resolved results: a developer-defined
  field's content is produced by the `resolveField` hook and refreshed via the
  right-click "Update Field" action; import preserves and export re-emits them.
- **Richer PDF export**: clickable internal links and a document outline
  (bookmarks pane) in exported PDFs.

### Fixed
- **TOC ribbon update preserves the existing look.** The ribbon "Insert / update
  table of contents" button no longer clobbers an existing (e.g. imported) TOC's
  title and per-level styling with the editor defaults when regenerating — it now
  preserves the current look, matching the "Update Field (TOC)" context action.
- **Import**: drop the orphan blank line left by a page-break-only paragraph.
- **Import**: correctly page-break a heading section that follows a continuous
  section break.

## [0.5.0] — 2026-06-15

### Added
- **Live TOC fields**: imported tables of contents round-trip as live OOXML fields
  on editor export.
- **TOC recalculation & headless generation**: patch-in-place page-number recalc
  and headless TOC generation as live OOXML fields, plus a backend route for
  TOC page-number recalculation.
- **Configurable webhooks**: the admin dashboard can subscribe webhooks to specific
  events, including `comment.mention`.

### Fixed
- Export table cell margins (`w:tcMar`) so they round-trip.
- Restore the `@forevka/wordcanvas/import` and `@forevka/wordcanvas/export`
  package subpath exports.

## [0.4.1] — 2026-06-14

### Added
- **Track changes & comments (review layer)**: a three-way Editing / Suggesting /
  Viewing mode switch. Suggesting mode records non-destructive, attributed
  insert/delete/format changes; comments thread on a range with reply/resolve and a
  docked Review pane; everything syncs in real time and rehydrates on join.
- **@-mentions in comments** from an embedder-supplied roster, with a
  `comment.mention` webhook.
- **View-only mode** via `new WordCanvas({ readonly: true })`.

## [0.3.2] — 2026-06-12

### Changed
- Package keywords/description and README position the editor as an open-source
  Syncfusion / OnlyOffice / DevExpress alternative.
- Deploy script: host/domain are parameterized via env vars (no hard-coded VPS).

### Fixed
- Resolve the `@forevka/wordcanvas/builder` subpath from `dist-lib` in the
  published `exports` map.

## [0.3.0] — 2026-06-11

### Added
- **Programmatic document builder** (`@forevka/wordcanvas/builder`) — a fluent
  API to compose documents in JS/TS, with `DocumentBuilder.fromTemplate(docx)` and
  an interactive playground example.
- **Multiple `WordCanvas` instances per page** (class-scoped chrome, per-instance
  teardown) and an `embed-multi` example served at `/multi`.

### Fixed
- Justify text beside floated images and stop tables from overlapping floats.
- Use a relative base for the library build so the import/export workers load
  correctly when the bundle is mounted under a subpath.
- Pin click-to-caret hit-testing to the clicked table cell.

## [0.2.0] — 2026-06-10

Initial public release of **`@forevka/wordcanvas`** — the canvas-rendered,
embeddable Word-style document editor extracted into a self-contained,
zero-runtime-dependency npm package, with standalone embedding examples served
from the web edge.

Editor capabilities at first release (built prior to packaging — see the
implementation history in [README.md](./README.md)):

- Canvas rendering pipeline with deterministic pretext line breaking and
  line-level pagination (widow/orphan, keep-with-next), page virtualization.
- Rich text & editing: typing via an IME proxy, justification, named styles,
  find & replace, hyperlinks, highlight, sub/superscript, format painter,
  autocorrect, soft breaks.
- Lists (bulleted/numbered, multilevel, inside table cells), tables (editable
  cells, rectangular selection, merge, borders/shading, row-level page breaks),
  images (insert/resize/square text-wrap), headers/footers with field tokens and
  story editing, sections & newspaper columns, fields/TOC, footnotes, bookmarks,
  hidden text.
- `.docx` and page-accurate PDF import/export (isomorphic browser/Node over
  bundled metric-clone fonts).
- Operational-transform collaboration over a WebSocket backend (gzip snapshot +
  parallel media upload on publish), an admin dashboard with document upload →
  docId, integration tokens for third-party `/upload`, and session webhooks.
- Mobile/touch input and a responsive ribbon.

[0.7.0]: https://github.com/Forevka/canvas-word/releases/tag/v0.7.0
[0.6.1]: https://github.com/Forevka/canvas-word/releases/tag/v0.6.1
[0.6.0]: https://github.com/Forevka/canvas-word/releases/tag/v0.6.0
[0.5.0]: https://github.com/Forevka/canvas-word/releases/tag/v0.5.0
[0.4.1]: https://github.com/Forevka/canvas-word/releases/tag/v0.4.1
[0.3.2]: https://github.com/Forevka/canvas-word/releases/tag/v0.3.2
[0.3.0]: https://github.com/Forevka/canvas-word/releases/tag/v0.3.0
[0.2.0]: https://github.com/Forevka/canvas-word/releases/tag/v0.2.0
