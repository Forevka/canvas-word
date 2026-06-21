# Changelog

All notable changes to **`@forevka/wordcanvas`** are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.7.3] — 2026-06-21

### Added
- **Customizable ribbon — `customizeRibbon` constructor option.** Embedders can now
  tailor the ribbon toolbar per editor instance, for custom macros, config popups,
  and informational popups. `customizeRibbon(api)` runs once at mount with a
  `RibbonApi` that addresses everything by **id**:
  - **Reorder / remove built-ins** — `moveItem` / `removeItem`, `moveGroup` /
    `removeGroup`, `moveTab` / `removeTab`, with `{ before }` / `{ after }` anchors.
    Every built-in tab/group/button has a stable namespaced id (`home`, `home.font`,
    `home.font.bold`, …); discover them at runtime with `api.tabs()` /
    `api.groups(tabId)` / `api.items(groupId)`.
  - **Add your own** — `addTab`, `addGroup`, and `addButton` (into built-in or custom
    groups). A button's `icon` accepts an SVG string, emoji, or text, with optional
    `active(fmt)` / `enabled(fmt)` state predicates synced live.
  - **Macro/popup context** — a custom button's `onClick` receives a
    `RibbonActionContext`: the editor handle (get/set document, export, mode, …) plus
    `getSelection()`, `insertText()`, `emit(name, payload)` (surfaced as a new
    `custom` event — `wc.on("custom", …)`), and `registerCleanup()`. The
    `getSelection()` / `insertText()` helpers are also available directly on the
    `WordCanvas` instance and `EditorHandle`. The `makeFloatingDialog` helper is now
    exported for building draggable, non-blocking config popups.

  Unknown ids are ignored with a console warning (configs survive editor upgrades),
  and omitting `customizeRibbon` leaves the ribbon exactly as before. New exported
  types: `CustomizeRibbon`, `RibbonApi`, `RibbonButtonSpec`, `RibbonActionContext`,
  `FloatingDialogOptions`, `DocSelection`. See `RIBBON.md`.
- **Theming & configuration via the constructor — `theme`, `overrideDefaultStyles`,
  and `behavior` options.** The editor's previously-hardcoded look and feel are now
  configurable per instance (multiple editors with different configs coexist on one
  page):
  - **`theme`** — pass a (partial) `EditorTheme` to recolor the editor chrome: the
    gray canvas/gutter, table gridlines, drawing-grid mesh, hyperlink & accent
    colors, formatting marks, TOC dot-leaders, footnote rule, image placeholder,
    column separators, the caret, find-highlight, review pins, page gap, and the
    rulers (`ruler: { bg, content, line, label, font }`). Omit any field to keep its
    built-in value. A ready-made **`darkCanvasTheme`** is exported to spread + tweak.
    Affects the on-screen editor only — exported PDFs keep the built-in look.
  - **`overrideDefaultStyles`** — override the LIBRARY's built-in default
    run/paragraph styles (body `fontFamily` / `fontSizePx` / `color` / `lineHeight`
    and `headingFontFamily`) for NEW/blank documents and the fallback stylesheet.
    A loaded `.docx` keeps its OWN defaults (`w:docDefaults` / `Normal`) — set
    defaults there, or here, not both.
  - **`behavior`** — tune `zoomStep` (default 1.1), the `zoomMin`/`zoomMax` clamp
    (0.25–5), `indentStepPx` (36), and the default drawing-grid `gridSpacingPx` (24;
    `view.gridSpacingPx` still wins).

  All three are fully back-compatible: omit them and the editor renders exactly as
  before. New exported types: `EditorTheme`, `DefaultStyleOverrides`, `EditorBehavior`.
- **Save to your own pipeline — `onSave` option + `exportDocx()` / `exportPdf()`
  handle methods.** Embedders no longer need the separate headless
  `@forevka/wordcanvas/export` entrypoint to wire up a Save button:
  - A new **`onSave`** constructor option. When set, the toolbar's **Export (PDF /
    DOCX)** buttons hand the produced file to your callback — a `Blob`, the same
    raw `bytes`, the `format`, and any export `warnings` (see `SaveEvent`) —
    instead of triggering a browser download. Return a promise to keep the UI
    responsive while you upload. Omit it to keep the default download behaviour.
  - New **`exportDocx()`** and **`exportPdf()`** methods on the `WordCanvas`
    instance (and its `EditorHandle`), each resolving to a `Blob`. Drive your own
    Save button and `POST` the result anywhere. Track changes are baked to the
    original baseline, exactly like the toolbar's Export. These work even when the
    ribbon is hidden (`view.toolbar: false` / `readonly`).

  Both paths reuse the editor's existing in-worker export, so there is no extra
  bundle cost and no `installMeasureHost()` dance.
- **Hide the built-in Export buttons** — new `view.exportPdf` and `view.exportDocx`
  flags (both default `true`). Set either to `false` to drop the corresponding
  File ▸ Export button — for embedders that ship their own export/save pipeline
  (via `onSave` or the `exportDocx()` / `exportPdf()` methods) and don't want the
  built-in download button. The whole Export group hides when both are off.

### Fixed
- **Resizing an image no longer drops the selection when the drag changes the page
  count.** Dragging an image's resize handle wide/tall enough to reflow content onto
  a new page (or back off one) used to abort the drag mid-stroke: the user had to
  release, re-select the image, and grab the handle again to keep going. The page
  layer rebuilt every page placeholder from scratch on any page-count change, which
  detached the selection frame — and the handle holding pointer capture — from the
  document, implicitly releasing the drag. Placeholders are now reconciled in place
  (surviving pages keep their DOM identity, so the captured handle stays attached),
  and the frame re-acquires pointer capture if a reflow does move the image to a
  different page mid-drag. Reconciling also avoids a full canvas/observer teardown on
  every reflow that adds or removes a page.
- **Published TypeScript types brought back in sync with the runtime API.** The
  hand-written `types/wordcanvas.d.ts` had drifted and was missing options and
  methods that already worked at runtime, so TypeScript embedders had to cast to
  reach them. Now fully mirrored: the `readonly` / `mode` / `allowedModes` /
  `knownUsers` / `view` / `theme` / `overrideDefaultStyles` / `behavior` constructor
  options; the `WordCanvasViewOptions`, `EditorTheme`, `EditorBehavior`,
  `DefaultStyleOverrides`, `EditMode`, `ReviewLayer` (+ `Suggestion` / `CommentThread`
  / `Comment` / `Fragment`) types and the `darkCanvasTheme` export; the review-layer
  methods (`getMode` / `setMode` / `getReview` / `getKnownUsers` / `setKnownUsers` /
  accept-reject suggestion(s) / `addComment` / `replyToComment` / `resolveThread`) on
  both `WordCanvas` and `EditorHandle`; and the `modeChanged` / `reviewChanged` events.

## [0.7.2] — 2026-06-19

### Added
- **Show/hide formatting marks.** A new **View ▸ Show** toggle (also surfaced on
  **Home ▸ Paragraph** as the ¶ button — both buttons share one state) overlays Word's
  non-printing marks on the canvas: a **center dot** for every space, a **→ arrow** in
  each tab gap, a **¶ pilcrow** at every paragraph end, and a **↵ arrow** at each manual
  line break. Settable on open via the `view` constructor option (`{ formattingMarks }`)
  and togglable through the editor handle (`setShowFormattingMarks` /
  `getShowFormattingMarks`). Presentational only — a paint overlay positioned from the
  existing layout geometry (so the marks land exactly where the caret would), with no
  relayout or document-model change.
- **Vertical ruler + drawing grid with snap-to-grid — for precise object placement.**
  A **vertical ruler** now runs down the left of the page (inch ticks + top/bottom
  margin shading), mirroring the horizontal one and tracking the topmost visible
  page as you scroll. Its top/bottom **margin boundaries are draggable handles**
  (Word's vertical-ruler resize): drag to set the page's top/bottom margins, committed
  as one undoable step via the same `applyPageSetup` path as the Page Layout dialog.
  A toggleable **drawing grid** paints a light gridline mesh on every page, and
  **snap-to-grid** snaps dragged anchored (floating) images to the grid so they land
  exactly on the lines — snapping the object's absolute page position, so it's correct
  for any anchor origin. New independent **View ▸ Show** controls — *Horizontal ruler*,
  *Vertical ruler*, *Show grid*, *Snap to grid*, and a **grid-spacing** selector
  (1/8″ / 1/4″ / 1/2″, default 1/4″) — plus a `view` constructor option
  (`{ ruler, verticalRuler, grid, snapToGrid, gridSpacingPx }`) to set the initial
  state. Presentational only — no document-model change (snapping reuses the existing
  anchor offsets).
- **Full initial-view control via the `view` constructor option.** Embedders can now
  set everything the reader sees on open, per mount: the **Outline/navigation**,
  **Bookmarks**, **Review** (track changes + comments), and **Activity** panels; the
  **ribbon toolbar** (hide it for a chromeless-but-editable surface, or start it
  **collapsed**); the **status bar**; the **initial zoom**; and the rulers/grid above.
  `view: { ruler, verticalRuler, grid, snapToGrid, gridSpacingPx, formattingMarks,
  outline, bookmarks, reviewPane, activity, toolbar, ribbonCollapsed, statusBar, zoom }`
  — every field optional, omit to keep its default. (The dev harness also accepts `?view=<json>` to
  preview a configuration.)
- **Flexible page layout — a redesigned Page Layout dialog.** The old preset-only
  Page Setup panel is replaced by a draggable, non-blocking dialog (Layout → **Page
  setup**) with a live, schematic scaled-page preview and an **inches ⇄ cm** toggle.
  Tabs cover everything that shapes a section's page, applied to the caret's section:
  - **Page** — Letter/A4/Legal presets *plus* custom width × height and a
    Portrait/Landscape toggle.
  - **Margins** — Normal/Narrow/Wide presets *plus* per-side numeric entry.
  - **Columns** — custom column count, equal *or* explicit per-column widths, custom
    spacing, and an optional **separator line** between columns (`w:cols/@w:sep`).
  - **Layout** — **header/footer band distances** (previously unsettable and dropped
    on export), page-number restart, and the Different-first-page / odd-&-even toggles.
  - **Background** — a **page color** (`w:background`) and a **page border**
    (`w:pgBorders`) with style, width, color, and measure-from page/text.
- **Page layout round-trips through DOCX and PDF.** Export now emits `w:orient`, real
  `w:header`/`w:footer` distances (fixing a bug that hardcoded both to 720 twips),
  `w:cols` separators and per-column `w:col` widths, `w:pgBorders`, and a document-level
  `w:background` (with `w:displayBackgroundShape` so Word paints it). All of it imports
  back, and the **PDF export renders page color, borders, and column separators
  identically to the on-screen canvas** (shared paint geometry — no drift).
- **Style constructors — a unified Style Manager.** Create, edit, and delete styles
  for every OOXML-styleable entity through one draggable, non-blocking dialog
  (Home → Styles → **Manage styles…**, or right-click → **Styles…**): a left rail
  lists styles grouped by kind, a center pane edits the selected style, and a right
  pane shows a **live preview painted by the real layout engine**. Covers:
  - **Paragraph styles** — full character + paragraph property editor (font, size,
    bold/italic/underline/strike, color, highlight, position; alignment, spacing,
    indents, keep-with-next/lines, page-break-before, outline level), with a
    cycle-safe *based-on* picker. Replaces the old `prompt()`-only "new style".
  - **Character styles** — now first-class: a `type` discriminator distinguishes
    them from paragraph styles, runs carry a character-style reference, and the
    ribbon gallery applies a character style to the selection (marked with `ⓐ`).
  - **List styles** — a per-level editor (format, bullet glyph, number pattern,
    indent, marker hang) with a live multi-level preview.
  - **Table styles** — a brand-new round-trippable style entity with **conditional
    formatting** (header/total row, first/last column, row/column banding, corner
    cells). Apply one from **right-click → Table Style ▸**; the effective per-cell
    fill/borders are baked onto the table and the live preview shows header shading
    and banding.
- **Styles always round-trip.** Import now preserves **every defined** paragraph,
  character, list, and table style (not just the ones in use), so a style you author
  survives save → reopen even before it's applied — matching Word. A new **"Show
  only styles in use"** filter (funnel button in the Home → Styles group) collapses
  the gallery to applied styles for documents that carry large style catalogs.
- **Merge duplicate styles.** A **Merge duplicates** action in the Style Manager
  collapses named styles that are identical except for their name (common after
  importing generated docs, or duplicating a style) into one survivor — content
  references and `basedOn` pointers are remapped, the default style is always kept,
  and it's a single undo. The button shows the count and disables when there are none.
- **Child documents (`WordCanvas.createChild()`).** A lightweight sibling document
  that shares the parent editor's *live* style context — stylesheet, list/numbering
  definitions, page section, and content-control/field maps — and renders or edits
  an arbitrary content slice (blocks, a fragment, runs, OOXML, or a named-style
  sample) with the **real layout engine + canvas painter** instead of an HTML
  approximation. Exposed publicly (`editor.createChild()` / `wc.createChild()`) and
  used internally to replace every in-app HTML-preview: the **Home → Styles gallery**
  swatches now show true `AaBbCc` samples in each style's resolved font/weight/size/
  color; the **content-control inspector** paints the real content (read-only) and
  hosts a **canvas-native editor** for editable controls that commits edited blocks
  straight back through the `replaceSdt*` commands (no `contentEditable` round-trip);
  and the **field constructor** previews the result in the document's real font. Each
  child owns its own layout engine, so its cache can't collide with the parent's.
- **Shared media in child documents.** A child proxies the parent's content-addressed
  media store: it renders the parent's images, and a child editor can **add new
  images** (`childEditor.insertImage(bytes, mime)`) whose bytes register back into the
  shared store — so they persist, export, and survive commit-back into the parent. The
  content-control inspector exposes this as an **Insert Image…** action for block-level
  rich-text controls.
- **Insert Image from your device.** The ribbon's Insert Image button now opens a file
  picker, registers the chosen image in the media store (content-addressed `mediaId`,
  so it persists and exports), and inserts it at its natural size (capped width) —
  replacing the previous fixed placeholder image.
- **Tracked structural edits (suggestion mode).** Paragraph split/merge, block
  add/remove, and table row/column add/remove are now recorded as **structural
  suggestions** instead of passing through untracked. The edit still applies to
  the live document; a `structural` record carries the applied op plus its exact
  inverse, so **reject** restores the original structure (re-merge a split,
  re-insert a removed block, etc.) and **accept** simply clears the record.
  Structural records show a block-level change-bar in the margin and are
  garbage-collected when their block no longer exists. `Accept all` / `Reject
  all` order structural records correctly relative to text changes.
- **Tracked paste & cross-paragraph delete (suggestion mode).** Multi-operation
  edits — a multi-block paste, or a deletion spanning paragraphs — are no longer
  passed through untracked. They are decomposed into their constituent text and
  structural changes, each tracked as a suggestion, and bundled under one group
  so the whole action accepts or rejects as a single unit. Rejecting a pasted
  block removes its text and reverses the paragraph splits; rejecting a
  cross-paragraph delete restores both the text and the paragraph break.

### Fixed
- **Table Borders & Shading now applies on Done.** Previously the dialog only
  applied when you clicked a preset/edge button; adjusting Width/Color/Style (or the
  fill) and clicking **Done** closed it without doing anything. Done now commits any
  spec/fill change you made without having clicked a preset (so "set Width = 5 →
  Done" borders the selection at 5px), while the live preset/edge buttons are
  unchanged.

### Changed
- **Editing dialogs are draggable, non-blocking floating panels.** The Borders &
  Shading, content-control inspector, field constructor, and TOC-options dialogs no
  longer dim the page or trap clicks — the document stays visible and interactive so
  edits are seen live — and each can be dragged by its header out of the way (shared
  `makeFloatingDialog` helper). Borders & Shading also gains clearer apply
  affordances (a hint + "Apply borders to" / "Individual edges" captions).

## [0.7.1] — 2026-06-16

### Added
- **Behind-text & in-front-of-text images (DOCX backgrounds).** Anchored drawings with
  `wrapNone`/`behindDoc` — e.g. a full-page decorative background — now import as
  absolutely-positioned images on a *behind-* or *in-front-of-text* layer instead of being
  forced into the text flow. They sit at their anchor offset (honoring `relativeFrom`
  page/margin/paragraph and negative offsets that bleed to the page edge), take **no flow
  height**, and don't reflow surrounding text. The engine paints layers in a fixed order
  (behind → text → front) that is the single source of truth for **both the canvas and the
  PDF renderer**, so they can't drift; pages with no anchored images are left untouched. New
  `ImageBlock.anchor` model field, and the anchor (incl. `relativeHeight` stacking order)
  round-trips back to `wp:anchor` on `.docx` export.
  - **Manipulation UI**: right-click an image for **Behind Text / In Front of Text** and
    **Bring to Front / Send to Back**; **click-drag** to reposition an anchored image; and
    resize a background up to the full page width. Gap-aware hit-testing makes a background
    selectable where text doesn't cover it (with a move cursor) while keeping the text on top
    clickable, and clicks outside the page deselect rather than snapping onto the background.
- **First-class field objects** — built-in Word fields are now real objects, not
  literal text. PAGE, NUMPAGES, DATE, TIME and IF carry a typed definition (parsed
  switches/arguments) and render as a result region **outlined in the editor like a
  content control** (gray field shading + a labelled tab). Right-click gives **Insert
  Field…**, **Edit Field…**, and **Update Field** (Word's F9): a visual **field
  constructor** picks the type and its format/switches/operands with a live preview.
  - Inline fields are marked per-run (`CharStyle.fieldId`, mirroring `sdtId`); the
    definition lives in `Document.fields` (now `kind: "builtin" | "custom"` + a typed
    `FieldSpec`). PAGE/NUMPAGES keep a `{page}`/`{pages}` token re-resolved per page;
    DATE/TIME/IF materialize their result. A new collab-safe `setField` op + a shared
    `@cw/shared` evaluator (`evaluateField`, `formatFieldDate`, `evaluateIf`).
  - **`.docx` round-trip**: built-in inline fields import as field objects and export
    back as real complex fields (instruction preserved verbatim). This also fixes the
    body/footer asymmetry — a body PAGE field is now a live field, not stale cached
    text. Untagged literal `{page}` text and existing header/footer behavior are
    unchanged.
- **Flagship showcase document** — the editor's initial state (no `docId`) is now a
  feature tour: a generated TOC, all field types (incl. inside table cells), every
  content-control kind, merged-cell and cross-page tables, images (block + square
  wrap), lists, footnotes, bookmarks, hidden text, and headers/footers.
- **Table-of-contents field options dialog.** Right-clicking a TOC now offers
  **Table of Contents options…** beside Update Field — a dialog to review and edit
  the field's switches (`\o` heading-level range, `\u`, `\h`, `\z`, `\p` separator)
  with a live field-instruction preview. Applying persists the instruction and
  regenerates the entries (a narrower `\o` range lists fewer levels), so it round-trips
  to `.docx`. Backed by a new collab-safe `setTocInstruction` op.
- **Programmatic builder now authors the full feature surface.** `@forevka/wordcanvas/builder`
  gained fluent methods for everything the model supports: inline **fields**
  (`pageField`/`numPagesField`/`dateField`/`timeField`/`ifField`/`customField`/`crossReference`),
  **content controls** (`contentControl` + per-kind sugar), **footnotes**, **bookmarks**
  (`bookmark`/`bookmarkRange`), super/subscript + letter-spacing + hidden text, a **table of
  contents** (`tableOfContents`, resolved from the document's headings at `build()`),
  **section breaks** with newspaper columns (`sectionBreak`/`columnBreak`), and richer styling —
  `defaultStyle`, custom `listDefinition`s, and reusable `tableStylePreset`s applied via
  `.table(rows, { style })`. No model changes — the builder mints the same structures the editor
  and exporter already consume, so it all round-trips to `.docx`. See
  [BUILDER.md](./BUILDER.md).
- **Headless PDF render — `POST /render.pdf`.** Stateless backend route: a raw `.docx`
  in, a rendered PDF out, for a producer that has no layout engine (e.g. a C# pipeline
  that emits Word fields it can't compute). It builds the table of contents from the
  document's headings, computes footer "Page X of Y" per page, then renders. Accepts
  multipart (`file` + optional `toc` JSON `TocOptions`) or a raw `.docx` body. Behavior
  is documented in
  [TOC-RENDERING.md](./frontend/src/export/pdf/TOC-RENDERING.md). The Swagger UI now
  exposes an **Authorize** control (integration API key via `X-API-Key`, or Bearer).

### Changed
- **Headless TOC generation now matches the source document.** The TOC builder moved
  into `@cw/shared` (shared by the editor and the headless render), honors the `TOC`
  field's `\o`/`\t` switches (level range, custom styles), and inherits the document's
  own TOC paragraph style — so an emitted `TOC \o "1-5"` reproduces the source's look
  with no per-level config. A TOC field that **already has entries is preserved** (only
  its page numbers are recomputed by layout); only an **empty** TOC field is built from
  the headings. Images and tables are never pulled into entries, even if mis-styled as a
  heading.
- `dev:online` now also launches the admin dashboard (port 5174, pinned via
  `strictPort`) alongside the backend and editor, so you can log in and mint integration
  tokens for the API.

### Removed
- The `POST /recalc.docx` and `POST /generate-toc.docx` backend routes, superseded by
  `POST /render.pdf` (the deliverable is a rendered PDF, not a field-baked `.docx`). The
  `@forevka/wordcanvas/recalc-docx`, `/generate-toc`, and `/recalc` library exports are
  unchanged.

### Fixed
- **A full-page background image no longer pushes the document onto a second page.** A
  `wrapNone`/`behindDoc` drawing (e.g. a decorative 8.5×11" page background) was laid out as
  a flow block, so it consumed the entire first page and shoved every paragraph onto page 2 —
  while Word keeps it on one page, behind the text. Such drawings are now out-of-flow
  background images (see Added), so the document paginates as authored.
- **Body PAGE/NUMPAGES fields now render their number, not the raw `{page}` token.**
  Inline page-number fields in body text (and table cells) are resolved against the
  final page map in a paint-only layout post-pass — the same path that already
  resolves TOC page numbers — so they show the page they land on and stay correct
  in both the on-screen canvas and the headless PDF render.
- **A bookmark in the same paragraph as an inline field no longer drops the field
  on export.** The `.docx` writer's bookmark-splicing path emitted runs one-by-one
  and bypassed the complex-field wrapping, so the field re-imported as plain text;
  it now groups field runs while still splicing bookmark markers at run boundaries.
- **List markers no longer draw on top of a floated image.** A numbered/bulleted
  list item flowing beside a square-wrapped (left/right floated) image had its text
  pushed clear of the image but its marker ("1.", "•") stranded at the margin, over
  the image. The marker now hangs off the first line's float-shifted start, tracking
  the text beside the float.
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

[0.7.3]: https://github.com/Forevka/canvas-word/releases/tag/v0.7.3
[0.7.2]: https://github.com/Forevka/canvas-word/releases/tag/v0.7.2
[0.7.1]: https://github.com/Forevka/canvas-word/releases/tag/v0.7.1
[0.7.0]: https://github.com/Forevka/canvas-word/releases/tag/v0.7.0
[0.6.1]: https://github.com/Forevka/canvas-word/releases/tag/v0.6.1
[0.6.0]: https://github.com/Forevka/canvas-word/releases/tag/v0.6.0
[0.5.0]: https://github.com/Forevka/canvas-word/releases/tag/v0.5.0
[0.4.1]: https://github.com/Forevka/canvas-word/releases/tag/v0.4.1
[0.3.2]: https://github.com/Forevka/canvas-word/releases/tag/v0.3.2
[0.3.0]: https://github.com/Forevka/canvas-word/releases/tag/v0.3.0
[0.2.0]: https://github.com/Forevka/canvas-word/releases/tag/v0.2.0
