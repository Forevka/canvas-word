# Changelog

All notable changes to **`@forevka/wordcanvas`** are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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

[0.6.0]: https://github.com/Forevka/canvas-word/releases/tag/v0.6.0
[0.5.0]: https://github.com/Forevka/canvas-word/releases/tag/v0.5.0
[0.4.1]: https://github.com/Forevka/canvas-word/releases/tag/v0.4.1
[0.3.2]: https://github.com/Forevka/canvas-word/releases/tag/v0.3.2
[0.3.0]: https://github.com/Forevka/canvas-word/releases/tag/v0.3.0
[0.2.0]: https://github.com/Forevka/canvas-word/releases/tag/v0.2.0
