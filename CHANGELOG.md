# Changelog

All notable changes to **`@forevka/wordcanvas`** are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- **Contextual floating toolbars — one framework, priority-based.** The two hand-wired floating
  mini-toolbars (image + text selection) are now unified under a single manager that shows the one
  most-relevant bar for whatever the caret/selection is on. Built-ins: the **image** bar (priority
  30), a new **table** bar (28 — merge / insert & delete rows & columns, shown when 2+ cells are
  selected), a new **hyperlink** bar (25 — Open / Edit / Copy / Remove, shown when the caret sits in
  a link), and the **text format** bar (20). Highest priority wins, so a selected image beats a text
  selection and a text *range* inside a link still shows the format bar. A new **public
  `contextToolbars` option** lets an embedder register their own context bars: each spec has a
  `when(ctx)` predicate over a `ToolbarContext` (format flags incl. `inTable`/`inContentControl`,
  selection, `hasRange`, `linkUrl`, anchor helpers), `buttons` (the same custom-button shape as
  `floatingToolbar` / the ribbon), an optional `priority` (default 15), and an optional `anchor(ctx)`.
  New editor accessors back it: `linkAtCaret()` (+ position-based `linkAtPosition`) and
  `getCellSelectionRect()`. See the [`context-toolbars`](./examples/context-toolbars) example.
- **Floating format toolbar (Word's selection mini-toolbar).** Selecting text now pops a compact
  toolbar just above the selection with the most-used character formatting — font family, font size
  (grow/shrink + presets), **Bold / Italic / Underline / Strikethrough**, text colour, highlight, and
  clear-formatting — so common edits don't require a trip to the ribbon. It anchors to the selection's
  start line, flips below when there's no room above, clamps to the viewport, tracks scroll/zoom, and
  hides on Escape, when the selection collapses, or when the selection scrolls out of view. Pressed
  state and the font/size readout mirror the ribbon. Edit-only (never shown in view-only mode) and
  mutually exclusive with the image mini-toolbar. **Fully configurable** via the `floatingToolbar`
  constructor option: pass `false` to hide it, or an object to set `enabled`, `onCaret` (also show at a
  bare caret, not only over a range), and `buttons` — pick and reorder the built-in controls (`"font"`,
  `"fontSize"`, `"bold"`, `"italic"`, `"underline"`, `"strikethrough"`, `"color"`, `"highlight"`,
  `"clearFormat"`, `"|"` separators) and add your own custom buttons (same `{ id, icon|label, tooltip,
  onClick, active }` shape and `RibbonActionContext` as a custom ribbon button). See the
  [`floating-toolbar`](./examples/floating-toolbar) example.
- **Reverse builder — `.docx` → DocumentBuilder code (`@forevka/wordcanvas/codegen`).** The inverse
  of `./builder`: `docxToBuilderCode(bytes)` (and the pure `emitBuilderCode(doc)`) generate editable
  TypeScript that calls the fluent `DocumentBuilder` API to reconstruct a document, so a doc authored
  visually becomes a **code template** a developer maintains and parameterizes. Runs reproduce exactly
  through the generic `.text(text, patch)` delta (same `resolveStyle` baseline the builder resolves),
  and stylesheet / list definitions / table styles reproduce verbatim. Fields the fluent API cannot
  yet express are **reported** in a structured `uncovered` list (node path + field + note) rather than
  dropped silently — the backlog for growing the builder. In the editor, develop mode adds an
  **Export builder code** button to the Developer ribbon tab. See [CODEGEN.md](./CODEGEN.md).
- **`ParagraphBuilder.listItem(listId, level)` — per-paragraph list membership.** The single-paragraph
  counterpart to the story-level `.list()`, for a list item that carries formatting `.list()` can't
  express (rich multi-run text, a named style, its own spacing/indent). Bridged in C#
  (`ParagraphBuilder.ListItem`). The first "frequently-used patch" the reverse builder's coverage
  report drove into the builder API.
- **Round-trip fidelity for Word's `w14`/`w15` identity metadata.** Five OOXML members that were
  previously dropped on import now survive a Word → edit → Word cycle:
  - **`w14:paraId` / `w14:textId`** — Word's persistent per-paragraph ids (what comment threads and
    co-authoring anchor to) are preserved on `Paragraph.paraId`/`.textId` and re-emitted. A source
    `<w:p>` that splits on a soft break gives its id to the first piece only, and export **de-dups**
    so a copy/pasted paragraph never emits a colliding id; new paragraphs emit none (Word assigns one
    on its next save).
  - **`w14:checkedState` / `w14:uncheckedState`** — a checkbox content control's exact checked/
    unchecked **glyph** (symbol font + code point) is preserved on `SdtProps.checkedSymbol`/
    `.uncheckedSymbol` instead of falling back to a default mark. Authorable via the builder
    (`.checkbox(true, { checkedSymbol, uncheckedSymbol })` / C# `Checkbox(…, checkedSymbol:, uncheckedSymbol:)`).
  - **`w15:docId`** — the document's persistent identity GUID (`settings.xml`) is preserved on
    `Document.docId` and re-emitted (declaring the `w15` namespace); the older `w14:docId` variant is
    also accepted on import.
- **Round-trip fidelity for Word's `wp14` drawing identity.** `wp14:anchorId` / `wp14:editId` on a
  floating/inline image's `wp:anchor`/`wp:inline` (the drawing-world analog of `w14:paraId`) are
  preserved on `ImageBlock.drawingId` and re-emitted (declaring the `wp14` namespace); export
  **de-dups** so a copy/pasted image never repeats an id.

### Fixed
- **Wrapping the first of two consecutive images no longer makes it vanish.** Turning an image into a
  square (text-wrap) float registers a float and does not advance the flow cursor, so following content
  flows beside it. But the layout engine's image-placement branch — unlike the table/equation/custom
  branches — did not drop below an active float first, so a float immediately followed by another
  **block image** placed the second image at the float's un-advanced `y`, painting it directly over the
  floated image and hiding it ([#201](https://github.com/Forevka/canvas-word/issues/201)). A block image
  is atomic and cannot flow beside a float, so it now drops below any active float before placing —
  matching how tables, equations, and custom blocks already behave.
- **Inline images no longer gain an unclickable vertical gap in Word.** An exported inline image lived
  in a `w:p` whose `w:pPr` carried only `w:jc`, so in Word the image paragraph inherited the default
  paragraph spacing from `Normal` (space-after plus a line multiplier — e.g. 1.5×). That opened a
  vertical gap around the image (most visible between the two consecutive images in the showcase's
  "Miscellaneous OOXML" section) that the editor never shows, since it lays an image out at exactly its
  pixel height with no paragraph spacing ([#198](https://github.com/Forevka/canvas-word/issues/198)).
  The image paragraph is now pinned to zero before/after and single (`w:line="240"`, `w:lineRule="auto"`)
  line spacing, so Word renders it flush — matching the canvas and the spacing-less `ImageBlock` model.
- **Exported `.docx` files now open in Microsoft Word.** Word validates every part against the strict
  OOXML `xsd:sequence` and rejects the whole document on the first violation; lenient consumers (Google
  Docs, LibreOffice) and our own order-independent importer did not, so several latent violations shipped
  unnoticed ([#180](https://github.com/Forevka/canvas-word/issues/180)). Three distinct classes were fixed:
  - **Element ordering** ([#191](https://github.com/Forevka/canvas-word/issues/191)). Children of
    `w:pPr`, `w:rPr`, `w:tcPr`, `w:sectPr`, `w:sdtPr` and `settings.xml` were emitted grouped by concern
    rather than in schema order (e.g. `w:tabs` after `w:spacing`, `w:shd` before `w:tcBorders`, `w:type`
    last in `w:sectPr`, `w:defaultTabStop` after `w:evenAndOddHeadersAndFooters`, content-control `w:lock`
    after the type element). A new schema-order normalizer sorts each property element's children to the
    canonical ECMA-376 sequence before emitting; a latent duplicate `w:bidi` in `w:pPr` is also gone.
  - **Display equations** ([#193](https://github.com/Forevka/canvas-word/issues/193)). A block equation
    was written as a bare `m:oMathPara` directly under `w:body`, which the schema forbids. It is now
    wrapped in a `w:p` (how Word stores it); the importer reads an `m:oMathPara` inside a `w:p` back to a
    display equation, so it round-trips.
  - **Colors** ([#193](https://github.com/Forevka/canvas-word/issues/193)). A 3-digit CSS hex color such
    as `#fff` was written as the invalid `w:val="fff"`; it now expands to the required 6 digits
    (`ffffff`). Applies to every color attribute (run color, shading fill, borders, page color).
  Export is now validated against the real schema (Open XML SDK) on the showcase document. No editor
  feature is stripped — this is correct OOXML for every document, so no "Word compatibility mode" is needed.
- **`wp14` percentage-positioned floating images no longer snap to the top-left corner on import.**
  Word wraps a floating object's percentage position (`wp14:pctPosHOffset`/`pctPosVOffset`) in an
  `mc:AlternateContent`, pushing the absolute `wp:posOffset` down into the `mc:Fallback` — so it is no
  longer a direct child of `wp:positionH`/`positionV`. The offset reader only looked at direct
  children, read nothing, and defaulted the position to `0`. It now recovers the `mc:Fallback`
  absolute offset (the percentage itself is not yet modeled), so a behind/in-front-of-text float lands
  where Word placed it.

## [0.10.2] — 2026-07-19

### Added
- **`handle.setCustomBlockData(id, data)` — update a custom block's data + re-render.** An undoable
  op that replaces a registered custom block's JSON `data` and bumps only that block's revision (no
  full-document rebuild). It's the ergonomic primitive for **asynchronous data** — `measure`/`paint`
  are synchronous (layout/paint run once per frame), so you fetch in your own async code and feed the
  result into the block: insert it in a loading state, then `setCustomBlockData(id, { state: "ready",
  … })` when the data arrives. Previously this required `setDocument`, which resets the whole document
  (drops undo/scroll). Backed by a new `setCustomBlockData` op/command (mirrors `setEquation`, with an
  inverse for undo). `examples/custom-block/` gains an “Insert async chart (loading → data)” button
  demonstrating the flow.

## [0.10.1] — 2026-07-19

### Fixed
- **PDF export no longer crashes on the browser main thread / StackBlitz.** Exporting a document to
  PDF failed with `TypeError: Cannot set property document of #<Window> which has only a getter`
  whenever the measure host ran on a thread with a real `document`: the inline main-thread path taken
  for PDF export of a **registered custom block** (see 0.10.0), and the StackBlitz WebContainer
  worker (whose export-worker global is `Window`-like). `installMeasureHost` assigned its DOM-free
  stub to `globalThis.document` unconditionally, but `window.document` is a read-only accessor with no
  setter. It now tolerates a read-only `document` (keeping the real one) while still routing pretext's
  width measurement through the fontkit `OffscreenCanvas` shim, so exported metrics are unchanged.
  Regression test simulates the getter-only `document`.

## [0.10.0] — 2026-07-19

### Added
- **Custom blocks render in PDF — a Canvas2D→pdfkit vector shim.** A registered custom block's
  `paint(ctx, box, data)` now renders into PDF (previously it reserved space + warned). The SAME
  `paint` is replayed through a pure-JS shim (`export/pdf/canvasToPdf.ts`) that implements the common
  Canvas2D surface — paths, `fillRect`/`strokeRect`, `arc`/`roundRect`/beziers, `fill`/`stroke`/
  `clip`, `fillText`/`measureText`, gradients, transforms, `globalAlpha` — and forwards each call to
  pdfkit as **crisp vector** ops. Being pure JS (no native canvas, no WASM) it runs **headlessly**:
  the Node backend and the bare-V8 ClearScript host render custom blocks with no extra wiring. In the
  browser editor, whose PDF export runs in a worker with no registry, a document containing a
  registered custom block is exported **inline on the main thread** so the block can paint. A block
  whose `paint` needs ops the shim can't translate (shadows, filters, pattern fills) can set
  `pdf: "raster"` — a reserved seam for a future headless-canvas (CanvasKit) rasterizer; until it
  ships, a `"raster"` block reserves its space and warns. Embedder `paint()` is sandboxed (clipped to
  the box, save/restore-balanced, errors caught) so a plugin bug can't corrupt the page. Real-pdfkit
  integration tests + shim unit tests.
- **Public decoration/overlay API at document coordinates (#175).** New `EditorHandle`
  methods — `setDecorations(specs)`, `clearDecorations()`, `invalidateDecorations()` — let an
  embedder draw custom paint-only overlays anchored to document positions: `highlight` (fill under
  text), `underline`, `box`, and `badge` (a marker at a `DocPosition`). Specs are described in
  document space (a `DocSelection` range or a `DocPosition`); the editor resolves them against the
  live layout via `geometry.selectionRects`/`caretRect` and re-resolves after every edit, so they
  stay anchored and follow reflow — and, like the track-changes overlays, it never measures text or
  re-breaks lines (paint-never-measures preserved). The resolver lives in the DOM-free
  `decorations.ts` with injectable geometry (unit-tested); the renderer gains one `setDecorations`
  scheduler method + two paint passes (under-text fills, over-text lines/boxes/badges) mirroring the
  review overlay feed. Decorations can be **interactive**: a spec with an `onClick` intercepts the
  click (instead of placing a caret) and shows a pointer cursor on hover, like a comment pin — the
  renderer hit-tests them via `decorationAt` and the input layer dispatches to the source spec. New
  `examples/decorations/` demonstrates it.
- **Public block-type registry — `registerBlockType` (#176).** Add a NEW document block type via a
  single `{ measure, paint, toOOXML? }` registration instead of hand-editing the ~8 dispatch sites
  that enumerate the built-in block union. A custom block is **canvas-drawn** and **atomic**: it
  measures to one box and paginates like an image/equation, lays out anywhere blocks go (body, table
  cells, header/footer bands), and its `data` is plain JSON so snapshot serialize/paste is free
  (deep-cloned on paste). It's a **first-class editor object** — click to select (a plain,
  non-resizable frame), Delete or right-click ▸ Delete to remove, with full undo/redo (wired through
  `hitTestCustom`/`objectRect` and a generalized `removeBlockObject`). A block in the model is
  `{ kind: "custom", customType, data, id, revision }` (new `CustomBlock`); the registered
  `paint(ctx, box, data)` gets a canvas context translated + clipped to the block's box. Wired
  through the engine (body + table-cell + band measure/placement, atomic gap logic), the layout tree
  (`PlacedBlock.custom`), the canvas painter, and both exporters — with the embedder `measure()`,
  `paint()`, and `toOOXML()` calls each guarded so a plugin bug can't crash layout/paint/export.
  A custom block has no native OOXML, so `.docx` export is lossy unless the type supplies
  `toOOXML(data)` — otherwise it emits a placeholder paragraph with a `custom-block-dropped` warning
  (PDF reserves the height and warns `custom-block-not-rendered`). An unregistered `customType`
  paints a visible dashed placeholder. New `examples/custom-block/` (a bar-chart block). By design a
  custom block's content is drawn (no internal caret) and its size is `measure()`-driven (no
  interactive resize); it never originates from an imported `.docx`.
- **Public command + keymap registry — the `commands` option (#174).** An embedder can now
  register named commands and bind keyboard shortcuts to them without forking the core — the
  additive counterpart to `customizeRibbon`. Each `EditorCommand` is `{ id, label?, keybinding?,
  run(ctx) }`; the handler `ctx` is the *same* `CommandContext` (== `RibbonActionContext`) a custom
  ribbon button gets, so a macro is interchangeable between a keystroke, a button, and the new
  `handle.runCommand(id)`. Keybindings use `Mod+Shift+K` syntax (`Mod` = Ctrl on Windows/Linux, ⌘ on
  macOS, so one binding is cross-platform); the built-in editing chords (Ctrl+B/I/U/Z/Y, Ctrl+Enter)
  always take precedence, and duplicate ids / conflicting chords are dropped with a `console.warn`
  (first registration wins). Pure chord parsing/matching/conflict logic lives in the DOM-free
  `commands.ts` (unit-tested like `ribbon.ts`); wiring is one `keydown` listener on the editor
  container, tied to teardown. New `examples/command-registry/` demonstrates it end-to-end.
- **Linked ("Link to File") images — `a:blip r:link` (#172).** Images whose bytes live
  OUTSIDE the document (a `TargetMode="External"` relationship targeting a URL) were dropped on
  import: the DrawingML parser only read `r:embed`, so a document whose images are all external
  (e.g. S3 URLs) loaded with no images at all. `parseDrawing` now falls back to `r:link` and stamps
  the URL onto a new `ImageBlock.externalSrc`; the editor displays it (the canvas loads the URL
  directly), export re-emits `a:blip r:link` + an External relationship instead of packing bytes,
  and `externalSrc` survives serialize (which blanks the runtime `src`) so it rehydrates on reload.
  Authorable via the builder (`.image(url, { linked: true })`) and C# (`ImageOptions.Linked`), with
  round-trip/serialize tests, showcase coverage, and `OOXML_COVERAGE.md` updated.
- **Host-owned resolver to embed linked images in headless PDF (C# ClearScript).** The bare-V8 host
  has no `fetch`, so linked images render as placeholder boxes in headless PDF. `WordDocument`
  gains `LinkedImageUrls()` and `ExportPdf(LinkedImageResolver)` / async / `Stream` overloads: the
  host supplies the bytes for each external URL with its own `HttpClient` (proxy / headers / TLS /
  timeout) and its own concurrency cap — the engine never fetches. DOCX export is unaffected (it
  re-emits `r:link`). A `linkedimg` benchmark verifies the resolver-fed PDF is byte-identical to the
  same image embedded directly.
- **Export warning when a linked image can't be embedded due to CORS.** The browser lets the editor
  *display* a cross-origin image (drawing needs no CORS) but blocks *reading* its bytes for export
  unless the host sends `Access-Control-Allow-Origin`. Client-side PDF/DOCX export now reports an
  `image-external-unfetchable` warning (in `ExportResult.warnings`, so embedders and the `onSave`
  hook see it) listing how many linked images couldn't be fetched, and the editor raises a visible
  toast — so an image that shows on screen but exports as a gray box is no longer a silent surprise.
  The fix is on the image host (add the origin to its CORS allowlist), not the library.
- **Round-trip fidelity for run proofing language `w:lang` (issue #168).** `w:lang`
  (`@w:val`/`@w:eastAsia`/`@w:bidi`) — which Word emits on nearly every run and in
  `docDefaults`/styles, making it the highest-frequency previously-dropped property — now
  round-trips as `CharStyle.lang` (a `RunLang` value object). Decoded on runs, character
  styles, and `w:rPrDefault`, cascaded through the style resolver (docDefaults bake onto each
  run), and re-emitted via the shared run serializer; runs with different language tags no
  longer merge. Round-trip-only — no layout/paint effect (in Word it drives spell-check,
  hyphenation, and CJK font resolution). Exposed on the `DocumentBuilder` (`.lang({ val,
  eastAsia, bidi })`) with a C# binding mirror, sampleDoc, showcase, and round-trip/decode/
  cascade tests. `OOXML_COVERAGE.md` updated.
- **Round-trip fidelity for the remaining East-Asian `w:pPr` toggles (issue #167).** `w:wordWrap`,
  `w:topLinePunct`, `w:autoSpaceDE`, and `w:autoSpaceDN` — the East-Asian paragraph toggles that sit
  beside `w:kinsoku`/`w:overflowPunct` in the schema and that #161 left out — were unmodeled (dropped
  on import, never re-emitted). They now round-trip as optional `ParaStyle` booleans via the same
  on/off seam (explicit `w:val="0"` preserved so an OFF can override an inherited `true`), with
  `DocumentBuilder` methods (`.wordWrap()`/`.topLinePunct()`/`.autoSpaceDE()`/`.autoSpaceDN()`), C#
  binding mirrors, showcase, and round-trip/decode tests. Round-trip-only — no layout behavior. Also
  adds `OOXML_COVERAGE.md`, a living tree of which WordprocessingML elements/attributes the
  import/export pipeline covers.
- **Round-trip fidelity for four low-frequency CJK / hyphenation toggles (issue #161).** `w:snapToGrid`
  (run + paragraph), `w:suppressAutoHyphens`, `w:kinsoku`, and `w:overflowPunct` (paragraph) were
  previously unmodeled — silently dropped on import and never re-emitted. They now round-trip as
  optional booleans on `ParaStyle` (and `CharStyle.snapToGrid` for runs), decoded via the shared
  on/off path (an explicit `w:val="0"` survives as `false` so it can override an inherited `true`) and
  re-emitted whenever defined. This is pure byte-fidelity preservation with no layout behavior —
  matching the issue-#62 minor-props pattern (`w:mirrorIndents`, `w:adjustRightInd`,
  `w:suppressLineNumbers`); the engine's existing kinsoku line-breaking is unchanged and unrelated to
  the editor's canvas "snap to grid" view aid. Exposed on the `DocumentBuilder`
  (`.snapToGrid()`/`.suppressAutoHyphens()`/`.kinsoku()`/`.overflowPunct()` on a paragraph, and
  `.effects({ snapToGrid })` on a run) with C# binding mirrors and showcase coverage.
- **"Organize Pages" — visual page reordering.** A new Layout → Pages button opens a slide-sorter
  overlay of live page thumbnails: drag a page to reorder, delete a page, double-click to jump to it,
  or focus a page and press `Ctrl`+Arrow to move it by keyboard (units are focusable, expose an
  `aria-label`, and get a focus ring; the delete button is reachable on focus). Enabled by default;
  hide it with the new `organizePages: false` WordCanvas option. Because pages are computed by layout
  (not stored like a PDF's), the movable unit is a page-break/section-delimited group, so reordering
  only re-sequences whole blocks — it never splits a paragraph or corrupts a section. Groups that
  render on the same page-run (e.g. a section whose tables spill onto the next group because a
  section break is swallowed by a hidden separator paragraph) merge into one unit and move together;
  a trailing section-break paragraph is left in place so other sections don't reflow; and a
  layout-verified pass pins a break only where a group actually merged, so no stray blank pages
  appear. Verified on a 43-page multi-section report: reordering keeps the page count stable, moves
  grouped content together, and doesn't split the Table of Contents.

### Changed
- **One dialog shell for the editor's floating dialogs (no visual change).** The Font, Paragraph,
  Borders & Shading, Page Layout, Field Constructor, Content-Control Inspector, Style Manager, and
  TOC Properties dialogs each hand-rolled the same backdrop + modal + header(title, ×) + body + foot
  DOM, the Escape/×-to-close wiring, the teardown `AbortController`, and the floating-panel setup.
  That scaffold now lives in one `createDialogShell` (`ui/dialogShell.ts`); each dialog keeps its own
  CSS (widths and backdrop shades intentionally differ per dialog) and its full public API, so the
  rendered DOM — class names, header order, badges, the Page Layout unit dropdown — is unchanged.
  Verified live: all eight dialogs mount, drag by the header, and close via Escape and ×.
- **Single-sourced OOXML mappings (no behavior change).** The border-edge → WordprocessingML
  encoding — previously copy-pasted five times across the docx exporter (cell/table borders,
  paragraph borders twice, run border, conditional table-style borders) — is one shared
  `borderEdgeXml` in the exporter's mappings module (the file that exists to prevent exactly this
  drift, which had already happened once). Word's 16 highlight colors were hand-maintained as two
  separate literal maps that had to stay mutual inverses (importer name→hex, exporter hex→name);
  the canonical map now lives in shared (`HIGHLIGHT_HEX`) with the inverse derived. The `w:shd`
  fill element and the `w:pBdr` emitter are shared instead of open-coded per site, and the
  paragraph-property import mapping is one `applyParaProps` core backing both the full-style and
  style-patch mappers (mirroring `applyRunProps`), so a new `w:pPr` field can no longer land in one
  and drift from the other.
- **DocumentBuilder: uniform warn-and-continue for invalid input.** `.effects({ widthScalePct })` /
  `.effects({ fitTextPx })` out of range and `.image()` without positive dimensions no longer THROW
  mid-chain — they record a warning (`effects-width-scale-invalid`, `effects-fit-text-invalid`,
  `image-size-invalid`) in `builder.warnings` and skip the invalid value/block, matching what
  `.withStyle()`/`.list()`/`.spacing()` always did. One contract for the whole fluent surface: a bad
  value can't abort a long chain, and every problem lands in the same diagnostics channel.
- **DocumentBuilder output is run-canonical.** Consecutive same-style `.text()` calls now coalesce
  into one run at author time. The merge requires the two styles to be STRUCTURALLY IDENTICAL —
  stricter than the editor's own merge criterion, which skips fields like `charStyleId`/`rtl` — so
  field/SDT/footnote/character-style/equation boundaries can never collapse. Builder output no
  longer violates the model's "adjacent equal-styled runs are merged" invariant until first edited.

### Fixed
- **DocumentBuilder diagnostics and provenance.** `.defaultStyle()` with an unknown id emits its own
  `default-style-missing:<id>` code — it previously shared `style-missing:<id>` with
  `.withStyle()`, and the dedup-by-code warning channel silently swallowed whichever fired second.
  `.withStyle()` now records explicit-key provenance on the runs it patches (like direct
  formatting), so a styled value that equals the resolved default survives table-style band baking.
  `.pageField()`/`.numPagesField()` accept a `style` option like every other field emitter (bold
  footer page numbers no longer need the raw `.field()` escape hatch). The default builder id seed
  is 10 random base36 chars (was 4 — ~1.7M values hit birthday-collision odds at a few thousand
  documents, which matters now that `mergeDocuments` folds builder outputs together).
- **Collab: suggestions/comments authored while briefly disconnected are no longer dropped.**
  Outbound review ops now queue while the socket is closed and drain on reconnect — the same
  guarantee core edits always had (receivers causal-hold on `dependsOnSeq`, so late delivery is
  safe).

### Performance
- **Session memory & bundle size.** The content-addressed media store no longer grows without bound:
  each editor mount registers a retention provider, and replacing the open document (docx open /
  `setDocument`) evicts bytes and revokes `blob:` URLs nothing references anymore — the undo-safe
  moment, since a replacement rebuilds the editor and its history. (With several concurrent mounts
  the store conservatively keeps everything, and destroying a mount doesn't evict — a remount may
  still rehydrate the same document.) The equation editor and symbol picker — click-driven dialogs
  that pulled the whole LaTeX toolchain (parser, serializer, symbol tables) into the initial editor
  chunk — now load lazily on first use, like the WebMCP polyfill and agent chat (~28 kB moved out of
  the critical path). Undo coalescing no longer re-copies the accumulated inverse-op array on every
  keystroke (O(n²) array churn across a typing burst): an open typing run appends amortized O(1) and
  folds back once when the run closes. Child-document editors (style previews) reuse the shared
  measuring engine instead of allocating a throwaway layout engine per mount.
- **Keystroke / repaint hot path — structural page diffing + font-string memoization.** `setTree`'s
  per-page change detection no longer `JSON.stringify`s every mounted page on every relayout (typing
  re-serialized 2–4 full pages per keystroke, and per frame during IME composition / column drags);
  it now compares pages structurally with identity short-circuits, so unchanged pages compare in
  O(blocks) via the engine's cache-shared line/table/band objects without descending into fragments —
  same over-inclusive "any difference repaints" safety. `charStyleToFont` — called per fragment on
  every layout, paint, and hit-test pass — is now memoized per (font registry, CharStyle object), and
  the canvas painter memoizes each run's `RunPaint` decision per style and skips redundant `ctx.font`
  assignments across same-styled fragments.
- **Pointer hot path — page-local hit-testing + one probe pass per frame.** Geometry queries
  (`hitTest`, `linkAt`, `pointOnText`, `inlineEquationAt`) scanned every line in the document per
  call; the line index is page-ordered, so they now binary-search the target page's contiguous span
  and touch only its lines — a click/hover on page 90 of a long report no longer walks pages 0–89
  (and the table-cell pin no longer walks the whole document even in table-free docs). Both hover
  `mousemove` handlers (resize/link/object affordances; content-control adornments + inspector) are
  coalesced to at most one probe pass per animation frame on the latest pointer position, so
  high-polling mice no longer run 6–8 geometry probes several times per frame. `locateParagraph` —
  probed per hover pass and per command dispatch — is now an O(1) lookup against a per-document
  location index (same WeakMap-on-identity contract as the paragraph index) instead of a full
  body + bands + notes walk.
- **DOCX import/export throughput on media- and text-heavy documents.** Importing a docx re-parsed
  the whole zip container once per distinct image (fflate's filtered `unzipSync` re-scans the central
  directory every call — O(images × entries)); media parts now inflate in ONE batched pass on the
  first media access, with the old per-part extraction kept as a fallback for containers with a
  corrupt media entry (a docx whose images are never resolved still pays nothing). Export no longer
  re-deflates already-compressed JPEG/PNG media bytes — `word/media/*` entries are stored at level 0
  (Word stores media essentially uncompressed too), while the XML parts keep default compression.
  `decodeRunProps`/`decodeParaProps` — the hottest import path, ~26 linear child scans per run bag —
  now index each property bag's children once and probe O(1).

### Added
- **C# `WordCanvasEnginePool` — safe engine reuse for multi-threaded hosts (ClearScript bindings).** A thread-safe,
  concurrency-bounded pool of `WordCanvasEngine` instances for ASP.NET Core / worker hosts, where the single-isolate
  engine must never be shared across threads. Register it as a singleton; `UseAsync`/`Use` lease one engine to one
  caller at a time (offloading the synchronous V8 pump to a thread-pool thread) and reuse engines across requests, so
  the bundle-load + font-install cost is paid once. `maxConcurrency` caps parallel work — bounding V8 heap and thread
  pressure at once — and the per-engine `WordCanvasEngineOptions` is fully configurable (bundle/fonts path, heap
  limit). Faulted engines are disposed rather than re-pooled; `PrewarmAsync` optionally pre-builds engines (bounded so
  it never exceeds `maxConcurrency` live engines). New README section documents the pattern.

## [0.9.0] - 2026-07-02

### Added
- **Merge / append documents — across TS/npm and C#.** A pure model
  `mergeDocuments(dest, source, opts)` / `mergeAll(docs, opts)` (shared) folds one document after another,
  reconciling every id space (blocks/cells, named + table styles, list definitions, content controls, fields,
  footnotes/endnotes, bookmarks) so the parts can't collide or alias; media dedupes by content hash. Style
  reconciliation has two modes (`useDestination` / `keepSource`); a section seam (`nextPage`/`evenPage`/`oddPage`
  keeps each part's own geometry + bands, `continuous`/`none` flow inline). Exposed on `@forevka/wordcanvas/query`
  (`mergeDocuments`/`mergeAll` + `DocumentEditor.append`, one undoable step via a new coarse `setDocument` op) and
  in C# (`WordDocument.Append(other, MergeOptions)`, `WordCanvasEngine.Merge(params…)`, unioning the embedded-image
  maps). Also `DocumentEditor.setSectionFooter`/`setSectionHeader`/`setSectionBand` (and C#
  `WordDocumentEditor.SetSectionFooter`/`SetSectionHeader` via a `StoryBuilder` callback) for a post-merge
  per-section footer pass. **Templating:** `DocumentEditor.replaceSdtContent(sdtId, source, opts)` /
  C# `WordDocument.ReplaceSdtContent` swaps a block-level content control's entire content with another
  document (reconciling all id spaces + media, preserving the control's ancestry) — find control "X" (e.g.
  C# `GetSdtsByTag`/`GetSdtsByAlias`, now also exposed) and fill it with a rendered section. New
  `examples/merge-docs` (TS/Vite) + `WordCanvas.Example.MergeReport` (C# — render
  parts in a loop, fold with explicit `MergeOptions`, per-section footers) demos + a table+logo footer recipe in
  the showcase; published
  `query.d.ts` (+ parity guard) and `builder.d.ts` (inline fields / bookmarks / footnotes on `ParagraphBuilder`)
  extended. See `MERGE_PLAN.md`.
- **C# `WordDocumentEditor.SetParagraphStyle(blockId, ParaStylePatch)` (ClearScript bindings).** Closes the last
  C#↔TS edit-parity gap — patch a paragraph's style (alignment, indents, spacing, breaks, outline level, direction,
  tab stops) from .NET, reusing the builder's existing `ParaStylePatch` record. Thin wrapper over the JS
  `setParagraphStyle` (no bundle change); the showcase patches the first paragraph.
- **C# query getters — fields / bookmarks / notes / lists / styles / location / text (ClearScript bindings).**
  Mirrors the TS read surface into .NET: `WordDocument.GetFields`/`GetField`/`GetFieldsByName`, `GetBookmarks`/
  `GetBookmark`, `GetFootnotes`/`GetEndnotes`, `GetListItems(listId)` (resolved markers), `GetStyles`/`GetStyleById`,
  `GetBlockPath(id)`, `PositionOfText(needle)`, `RangeText(startBlockId, startOffset, endBlockId, endOffset)`,
  `GetSdtValue(id)`, and `IndexOnPage(blockId)`. Backed by new `queryBridge` mappers (`queryFields`/`queryBookmarks`/
  … → flat DTO records); the C# showcase now prints field/style/bookmark/footnote counts and exercises
  position/range/value/block-path. Also adds a **regex overload** `ReplaceAllText(pattern, replacement, flags)`
  (via a `newRegExp` entry helper) so the .NET editor can reach the JS `RegExp` branch; the bridge parity guard now
  also validates `_engine.Api` calls made from `WordDocumentEditor`.
- **C# `WordDocumentEditor` completion — SDT value/unwrap + edit ergonomics (ClearScript bindings).** Mirrors the
  TS edit facade so the .NET write surface matches: `SetSdtValue` (dropDown/comboBox select), `RemoveSdt(id,
  deleteContents?)` (unwrap), and the ergonomic bulk/structural edits `ReplaceAllText`, `SetStyleByName`,
  `MoveBlock`, `InsertTableRowAt`, `DeleteColumnByHeader`. Thin wrappers over the existing JS `DocumentEditor`
  methods (no bundle change), covered by the bridge parity guard; the C# showcase now selects a dropdown value,
  find/replaces, and moves a block.
- **C#↔JS bridge parity guard + C# SDT nesting query completion (ClearScript bindings).** A new TS test
  (`csharpBridgeParity.test.ts`) scrapes the `InvokeMethod`/`GetProperty` names the C# bindings call
  (`WordDocumentEditor` → a real `DocumentEditor` member; `WordDocumentQuery` → a wired JS bridge fn) and fails
  `npm test` if a JS method is renamed/removed out from under the C# side (which CI otherwise wouldn't catch,
  since it doesn't build .NET). Also completes the C# SDT nesting surface with `WordDocument.GetSdtAncestors(id)`
  / `GetSdtDescendants(id)` (pure LINQ over the flattened `SdtInfo` list, no bridge round-trip).
- **`walkRuns` run-level traversal + nested-cell ancestry (`@cw/shared` + `@forevka/wordcanvas/query`).** A new
  `walkRuns(doc, visit)` primitive visits every run with its paragraph, run index, and full enclosing content-control
  chain (`RunContext`) — the run-level companion to `walk`. `BlockContext` also gains a `cellPath` (full `CellRef[]`
  ancestry outer→inner) that is present only for a block nested in a table-within-a-cell (≥2 cells deep), so
  single-cell blocks keep their existing `{ container, cell }` shape — additive, no breaking change.
- **Edit-facade ergonomics on `DocumentEditor` (`@cw/shared` + `@forevka/wordcanvas/query`).** Higher-level
  one-undo edits: `replaceAllText(pattern, replacement)` (find/replace across every paragraph — string replaces
  all, RegExp per run; matches spanning a style boundary are preserved), `setStyleByName(blockId, styleName)`
  (resolve a human style name → styleId and bake the resolved paragraph + run formatting plus the reference),
  `moveBlock(blockId, toIndex)` (reorder a top-level body block), and table edits `insertTableRowAt(tableId,
  rowIndex, cellTexts?)` / `deleteColumnByHeader(tableId, headerText)`.
- **Range / text addressing helpers (`@cw/shared` + `@forevka/wordcanvas/query`).** `rangeText(doc, selection)`
  returns the text a selection covers (a single-block range slices that block; a multi-block range spans top-level
  body blocks joined by newlines; endpoints auto-ordered), `positionOfText(doc, needle)` returns the first
  `DocPosition` of a substring in reading order (target an edit without hand-computing offsets), and the
  layout-backed `indexOnPage(pages, blockId)` reports a block's page index + order within a `getPages` map.
- **Document query getters: fields, bookmarks, notes, list items, styles, block location (`@cw/shared` + `@forevka/wordcanvas/query`).**
  Rounds out the read surface beyond paragraphs/tables/sections: `getField`/`getFields`/`getFieldsByName`/`getFieldBlocks`
  (custom + built-in fields and their result region), `getBookmark`/`getBookmarks`, `getFootnotes`/`getEndnotes`
  (note stories by ref id), `getListItems(listId)` (paragraphs bound to a list, in body reading order, each with its
  **resolved marker** — mirrors the layout engine's numbering pass, purely), `getStyles`/`getStyleById`, and
  `blockPath(id)` ("where is this block" — its container/cell/note context). Published on the npm `/query` subpath
  with hand-written types (`BookmarkEntry`/`NoteStory`/`ListItem` + `FieldDef`/`NamedStyle`/`BookmarkRange` re-exports),
  guarded by the parity check.
- **SDT value select + unwrap (`@cw/shared` + `@forevka/wordcanvas/query`).** Rounds out the content-control edit
  surface: `getSdtValue(doc, id)` reads a control's value in the shape its `type` implies (`text`, plus `checked`
  for checkboxes and the resolved `selected` value for dropDown/comboBox); `DocumentEditor.setSdtValue(id, value)`
  selects a dropDown/comboBox option (by listItem value then display; comboBox allows free text; dropDown requires
  a listed option); and `DocumentEditor.removeSdt(id, { deleteContents? })` removes a control by **unwrapping** it —
  stripping its id from every member run and body block path so **nested controls survive**, then deleting its
  props (keeping the "every path id has a props entry" invariant), optionally deleting the wrapped content too.
- **C# SDT (content control) query + edit parity (ClearScript bindings).** The templating surface now reaches
  .NET. Query: `WordDocument.GetSdts()` returns the whole control forest flattened — each `SdtInfo` carries its
  type/tag/alias/checked/placeholder, its nesting links (`ParentId`/`ChildIds`/`Path`/`Depth`), and the text it
  encloses — plus `GetSdt(id)`, `GetSdtRoots()`, `GetSdtChildren(id)`. Edit: `WordDocumentEditor.SetSdtText(id, text)`
  (fill a field, nesting-preserving), `SetCheckbox(id, checked)`, and `SetSdtProps(id, SdtPropsPatch)` (tag/alias/
  placeholder/date-format/locks; `type` stays fixed). Backed by a `querySdts` JS bridge fn; the C# showcase now
  lists the document's controls and fills/toggles them.
- **SDT (content control) editing on `DocumentEditor` — the write half of templating (`@cw/shared` + `@forevka/wordcanvas/query`).**
  Fill and update content controls headlessly, over the same op engine (undo/redo for free): `setSdtProps(id, patch)`
  (merge alias/tag/checked/list/locks, preserving `type`), `setCheckbox(id, checked)` (checkbox controls), and
  `setSdtText(id, text)` — the killer "fill this field" primitive: replaces the text of a control occupying a single
  paragraph (inline or block-level), preserving the control's ancestry so **nested controls survive**, and clearing
  any placeholder flag in the same undoable step. Multi-block controls throw (edit those by block id). Published on
  the npm `/query` subpath and guarded by the parity check.
- **SDT (content control) query API — the primary templating surface (`@cw/shared` + `@forevka/wordcanvas/query`).**
  Content controls are how documents get templated, so the query layer now treats them first-class, including
  **nested controls** (an SDT inside another — membership is an ordered `sdtPath` ancestry, so nesting is native).
  Flat lookup: `getSdt(id)`, `getSdts()`, `getSdtsByTag(tag)`, `getSdtsByAlias(alias)`. Nesting tree:
  `getSdtNodes()` (each node carries `parentId`/`childIds`/`path`/`depth`), `getSdtRoots()`, `getSdtChildren(id)`,
  `getSdtAncestors(id)`, `getSdtDescendants(id)`. Content: `getSdtBlocks(id)` (block-level members) and
  `sdtText(id)` (the enclosed text — the "read the value" half of a template round-trip, covering block-level
  and inline membership, nested controls included). Exposed on the npm `/query` subpath with hand-written
  types (`SdtMatch`/`SdtNode`, guarded by the compile-time parity check).
- **C# in-place edit surface `WordDocumentEditor` (ClearScript bindings).** `WordDocument.Edit()` opens a
  stateful editor over a JS `DocumentEditor` in V8 (the write half of WordprocessingDocument-style access):
  `SetParagraphText`, `InsertText`/`DeleteText`/`ReplaceText`, `RemoveBlock`, `InsertParagraphAfter`/
  `InsertParagraphBefore`, and `Undo`/`Redo`/`CanUndo`/`CanRedo`. `ToDocument()` returns a handle over the
  edited model (preserving embedded image bytes) to query or export. The C# showcase now edits, undoes,
  redoes, and round-trips a change through docx.
- **C# read-only query surface on `WordDocument` (ClearScript bindings).** An imported or built document
  can now be inspected from .NET — the read half of WordprocessingDocument-style access: `GetParagraphs()`
  (every paragraph with its container/table-cell/note location, style name, outline level), `FindText(needle)`,
  `GetSections()` (per-section page geometry + block range), and `GetPages()` (runs a layout pass in V8 and
  reports the block ids on each page — "what's on page N"). Backed by a JS query bridge on the ClearScript
  entry; results marshal into `ParagraphInfo`/`SectionInfo`/`PageInfo` records. The C# showcase now imports
  its own exported docx and prints a query summary.
- **Page query + public `@forevka/wordcanvas/query` subpath.** A new package export exposes the document
  query + edit API to embedders: the traversal/find helpers, `DocumentEditor`, section enumeration, and a
  new layout-backed **`getPages`** / `pageOfBlock` — the answer to "what's on page N" (pages don't exist in
  the model; `getPages` runs a layout pass and returns a serializable per-page map of placed block ids +
  geometry, honoring `pageNumberStart`). Hand-written self-contained types ship in `types/query.d.ts`.
- **Document edit facade `DocumentEditor` (`@cw/shared`).** A headless, ergonomic editing layer over the
  `applyOp` operation engine — the rough analog of mutating a .NET `WordprocessingDocument` and saving.
  Holds a mutable `doc` (every edit swaps in a new immutable value via structural sharing), translates
  high-level calls into the existing typed ops, and keeps an undo/redo stack for free from the engine's
  inverses. Methods: `setParagraphText`, `insertText`/`deleteText`/`replaceText` (UTF-16 offsets),
  `setParagraphStyle`, `insertParagraph` (before/after a top-level block; clones the reference style
  minus structural markers and mints an id), `removeBlock`, plus `undo`/`redo`/`canUndo`/`canRedo` and
  `find`/`getParagraph` conveniences. Not the interactive editor's Command/caret machinery — plain data,
  usable from Node, the browser, and the C# bindings.
- **Document query API (`@cw/shared`).** A read-only traversal layer over the document model, the
  rough analog of .NET's `WordprocessingDocument` descendant queries: `walk` (visit every block,
  descending into table cells, header/footer bands, and note bodies), `getParagraphs`/`getTables`/
  `getImages`, `findParagraphs` (substring or RegExp, reporting each match's container/cell/note
  context), `getBlockById` with typed `getParagraphById`/`getTableById`/`getImageById` narrowers,
  `textOf` (block plain text; tables join cells with tabs, rows with newlines), and `getSections`
  (per-section page geometry with the top-level block range each covers). Pure and DOM-free; page-level
  queries are not here (pages exist only after layout).

### Changed
- **`effectiveSection`/`resolveSections` moved into the shared model core** (`@cw/shared`) from the
  layout engine, so the editor, the exporter, and the new query API resolve sections identically. The
  layout engine re-exports them, so existing importers are unaffected; `resolveSections` now also
  reports each section's `index` and `startBlock`.

## [0.8.1] — 2026-07-01

### Fixed
- **Inline bookmarks no longer drift or grow across a `.docx` open → save → open cycle.** A bookmark
  whose boundary fell in the middle of a run used to snap to the run's *end* on export (the writer
  only emitted bookmark markers between runs). Once import coalesced the bookmarked run with an
  adjacent same-style run, the bookmark expanded to swallow it — its span grew on every save until it
  hit the run boundary. The writer now splits a run at an interior bookmark offset so the marker lands
  on the exact character. Separately, the importer counted a footnote/endnote reference as zero-width
  when resolving bookmark offsets, but the model paints the reference's number (1+ chars), so a
  bookmark positioned after a note reference drifted one character early; the importer now shifts
  bookmark offsets past note-reference and inline-equation expansions. The default showcase document
  now survives repeated export→open→export as a stable fixed point.

## [0.8.0] — 2026-07-01

### Performance
- **Caret navigation on large documents is dramatically faster.** The document data-access
  helpers — `paragraphsOf`/`blockById`/`blockIndexOf`, body-vs-band membership, and the selection
  controller's caret-navigation paragraph list — used to re-walk the entire block tree (body +
  table cells + the six header/footer bands + footnotes/endnotes) and allocate a fresh array on
  *every* call, several times per keystroke. They are now memoized per immutable document identity
  via `WeakMap<Document, …>` caches (an id→paragraph index, an O(1) band-id set, and a cached
  navigable-paragraph list); because every edit returns a new `Document`, the cache auto-invalidates
  with no bookkeeping. On a ~3,300-paragraph appraisal report this cut the per-arrow-key data-access
  work from ~8 ms to ~0.03 ms (≈240×); editing and rendering share the same index. Behaviour is
  unchanged.

### Changed
- **Removed duplicated ribbon controls.** The Home ▸ Font hyperlink button was dropped (Insert ▸
  Links is the single, Word-canonical entry point for hyperlinks), and the two adjacent Home ▸
  Paragraph "shading" / "borders" buttons — which opened the identical Paragraph dialog covering
  both — were collapsed into one "Borders & shading" button.

### Fixed
- **Hebrew text no longer renders as tofu ("x") in PDF export.** Only CJK and Arabic had bundled
  fallback faces, so Hebrew runs fell through to a Latin metric-clone that lacks Hebrew glyphs and
  exported as `.notdef` boxes. A bundled **Noto Sans Hebrew** fallback (single Regular face, OFL 1.1)
  now joins the CJK/Arabic fallbacks: Hebrew runs are script-split onto it automatically so they
  measure, render, and subset-embed with a real face out of the box. On by default; opt out with
  `cjk: { hebrewFallbackFont: "" }`, or override with a registered custom family.
- **A cropped image that fails to decode no longer blanks the rest of the PDF page.** The PDF
  renderer clips to the crop window before drawing the image; if the image bytes were undecodable
  (e.g. an SVG, which pdfkit can't rasterize) the draw threw *after* the clip was applied but
  *before* it was restored, so the clip leaked and every following element on the page — body text,
  header, and footer — was clipped away and rendered invisible (white-on-white). The clip is now
  always restored (even on a decode error), so a bad image degrades to a placeholder box without
  affecting the rest of the page.
- **The showcase document's demo images now appear in PDF export.** They were SVG data URIs, which
  pdfkit can't decode, so they exported as gray placeholder boxes (and, for the cropped one,
  triggered the clip leak above). They are now a bundled raster (PNG) tile. General SVG support in
  the exporter is tracked in issue #116.
- **Arrow keys no longer stall next to a page-number field or hidden run.** With the caret just
  before a `PAGE`/`NUMPAGES` field (e.g. "Page 1 of 3") — or before a bookmarked run that follows
  hidden (`w:vanish`) text, as in the sample document — pressing ◀/▶ appeared to do nothing for
  several presses. Such runs occupy multiple model offsets but paint at a single point: a token
  field collapses its whole `{page}` range onto one resolved glyph (via `offsetMap`), and a hidden
  run lays out zero-width. Caret movement stepped one *model* offset per press, so it walked those
  dead offsets one at a time while the visible caret sat still. Horizontal movement now skips over
  offsets that resolve to the same painted caret point, so a single press always moves the caret
  past the field/hidden run (matching Word); ordinary text navigation is unchanged.
- **Line numbering no longer bleeds across a dropped section break on import.** A document with a
  line-numbered section (`w:lnNumType`) preceded by a plain, geometry-preserving Next Page break
  reopened with line numbers printed beside *every* line from page 1 — the whole document looked
  like it had become one numbered block. Line numbering (and page-number restart) are per-section
  OWN properties that never inherit, but the importer flowed the geometry-preserving break instead
  of keeping it, merging the unnumbered lead-in into the following numbered section so its property
  bled backward. Such a break is now preserved whenever the section it closes differs from the
  following section (the next `w:sectPr`, or the body `w:sectPr` for the last one) in line numbering
  or page-number restart; footer-only Next Page breaks with matching properties still flow.
- **Rows/columns added via the context menu now inherit the table's cell formatting.** Insert →
  Row Below / Column Right created cells with no borders, shading, or margin, so they fell back to
  the engine's bare defaults (light grid, no fill, default padding) instead of matching the table.
  Table-level defaults (`w:tblBorders`/`w:shd`/`w:tblCellMar`) are baked onto every cell at
  import/build time, so a new cell now copies the neighbouring cell's borders/shading/margin/vAlign/
  textDirection/noWrap/fitText/hideMark (content, merges, and preferred width are not copied).
- **The caret now follows the text angle inside vertical-text (`w:textDirection` tbRl/btLr) cells.**
  Previously the insertion caret in a rotated cell was drawn as an upright vertical bar — as if
  the text were horizontal. `caretRect` now reports the cell's ±90° rotation and the renderer
  rotates the caret bar about its center, so it sits horizontally across the column to match the
  rotated text. Horizontal cells are unaffected; remote-collaborator carets keep their upright bar
  (their name flag must stay upright).
- **Table column-resize grips now line up correctly on right-to-left (`w:bidiVisual`) tables.**
  Under `w:bidiVisual` the grid mirrors about its width (model column 0 paints at the right),
  but the column-boundary hit-test walked `colWidths` left-to-right from the table's left edge
  as if LTR — so the draggable grips appeared off the visible separators ("out of order") and
  mapped to the wrong model column, and the drag pushed the wrong way. The hit-test now mirrors
  the boundary x about the grid (keeping `boundaryIndex` in model order) and the drag flips its
  delta sign for bidiVisual. `w:tblInd`/alignment were never the cause (the offset is already
  folded into the table's rendered x); rows are unaffected by column mirroring.
- **PDF export: non-Latin symbols (✓ U+2713, ☒ U+2612, ballot boxes, dingbats) now render
  as real glyphs instead of `.notdef`/tofu.** The root cause was that the export resolved one
  bundled face per run with no per-glyph fallback; characters outside that face silently
  produced the missing-glyph box. The fix adds a **per-glyph fallback** in the EXPORT path
  only (on-screen path is unchanged): `glyphFallback.ts` (`segmentByFace`) splits any text
  string into face-homogeneous segments — ASCII stays in the fast path, non-ASCII code
  points that the primary face lacks are routed to **StixTwoMath** (the bundled math/symbol
  font, which covers the full dingbat and miscellaneous-symbols Unicode ranges including
  ✓ ☐ ☑ ☒). The split is applied in BOTH `fontkitContext.ts` (measurement) and
  `paintBlock.ts` `paintLine` (painting) so reserved widths always match painted glyph
  positions. Fallback faces are subset-embedded in the PDF (`renderPdf.ts`). NotoSansSC
  is intentionally excluded from the per-glyph fallback chain: CJK text is pre-routed
  by `scriptSplitRuns` and the `cjk.fallbackFont: ""` opt-out must not be bypassed.
  Latin-only output is byte-identical (the ASCII fast path is unchanged). Closes #104.

### Added
- **Arabic text in PDF export (issue #105).** Arabic text now renders with correct
  contextual joining forms (initial/medial/final/isolated letter shapes) and
  right-to-left visual order in PDF export instead of `.notdef`/tofu. A **Noto Sans
  Arabic** Regular face (~240 KB, SIL OFL 1.1) is bundled alongside the existing
  CJK (Noto Sans SC) fallback; `scriptSplitRuns` in `layout/prepareCache.ts` now
  splits runs at Arabic ↔ non-Arabic script boundaries (Unicode ranges U+0600–06FF,
  0750–077F, 08A0–08FF, FB50–FDFF, FE70–FEFF) and retargets those sub-runs to the
  Arabic face. fontkit's `layout()` applies GSUB contextual substitution (joining)
  when measuring and when pdfkit embeds the glyphs, so the export is both shaped and
  metrically consistent with the on-screen canvas render. The fallback is on by
  default (mirrors the CJK default-on approach); pass `cjk: { arabicFallbackFont: "" }`
  to opt out. The sampleDoc already exercises Arabic RTL paragraphs. Deferred:
  Hebrew and other RTL scripts (Syriac, Thaana) do not yet have a bundled fallback
  face; they will render correctly on systems that have a matching font registered
  with the browser, but will be tofu in PDF export until a corresponding subset is
  added.


- **Vertical cell text (`w:textDirection` `tbRl`/`btLr`) — real 90° rotation.** Cells
  whose text direction is `tbRl` (top→bottom, columns right→left) or `btLr` (bottom→top,
  columns left→right) now render rotated instead of flowing horizontally. The layout
  engine measures a rotated cell in a swapped frame — the laid-out text length drives the
  **row height** and the stack of line heights drives the **column width** — so a vertical
  header column auto-sizes narrow and tall (AutoFit), exactly like Word (`measureTable` /
  `placeTable` in `layout/engine.ts`, new `PlacedTableCell.rotation`). The canvas renderer
  (`paint/renderer.ts`) and PDF exporter (`export/pdf/paintBlock.ts`) wrap the cell's
  content in a matching `translate`+`rotate`, so canvas and PDF agree. Caret placement and
  click hit-testing inverse-rotate through the cell so clicking into a vertical cell lands
  the caret correctly (`layout/geometry.ts`). Round-trip and the existing UI are unchanged
  (the model/import/export already preserved the value). The East-Asian upright variants
  (`tbRlV`/`tbLrV`) degrade to the same 90° clockwise rotation without per-glyph
  uprighting; `lrTb`/`lrTbV` stay horizontal. Known limitations: the caret renders as a
  vertical bar (placement is correct, orientation is not) and per-grapheme selection inside
  a rotated cell is approximated by a containing box. A `tbRl`/`btLr` demo table is in the
  showcase (`model/sampleDoc.ts`).
- **Insert → Symbol (editor UI).** A new **Symbol** button on the Insert ribbon tab
  opens a floating symbol picker: choose a symbol font (Symbol, Wingdings, Wingdings 2,
  Wingdings 3, Webdings) and click a glyph from the Private-Use grid, with a
  "recently used" row persisted in `localStorage`. Picking inserts a run carrying the
  existing `CharStyle.symbol` marker (`{ font, char }`) with its `text` set to the
  decoded glyph and `fontFamily` set to the symbol font, via a new `insertSymbolCmd`
  in `editor/commands.ts` (mirrors the insert-equation/insert-field path, undoable).
  The model field and `w:sym` docx round-trip already shipped — this only adds the
  authoring UI.
- **Font dialog + caps / small-caps / double-strike ribbon toggles (editor UI).**
  The previously-disabled Home-tab "Text effects" launcher now opens a real, draggable
  **Font dialog** (`ui/fontDialog.ts`) that authors the run-level `CharStyle` fields that
  already round-trip but had no UI: all-caps (`w:caps`), small caps (`w:smallCaps`), double
  strikethrough (`w:dstrike`), underline **style** + **color** (`w:u/@w:val` + `@w:color`),
  baseline raise/lower (`w:position`), width scaling (`w:w`), character spacing, the kerning
  threshold (`w:kern`), the emphasis mark (`w:em`), the `outline`/`shadow`/`emboss`/`imprint`
  effects, and the fit-text width (`w:fitText`). Apply routes a single `Partial<CharStyle>`
  patch through `setCharStyle` (one undoable edit over the selection / pending typing-style).
  Caps, small caps, and double strikethrough also get quick toggles beside B/I/U (via
  `toggleCharStyle`), and `currentFormat()` now reflects all of these so the toggles light
  up and the dialog opens seeded from the caret. Editor-only — no model, import/export, or
  C# change.
- **Page Layout dialog — section start & line numbering controls.** The Page Layout
  dialog's Layout pane now exposes the section-start type (`w:sectPr/w:type`: New page /
  Even page / Odd page → `SectionPatch`/`SectionProps.breakType`) and a line-numbering
  group (on/off, count-by, start-at, restart per page/section/continuous, and distance
  from the text edge → `lineNumbering`). Both seed from the caret's section and pack
  into `applyPageSetup`, which writes them onto the terminating section-break paragraph
  or `doc.section` with full undo. The model fields and `.docx` round-trip already
  shipped (#82); this wires the editor UI to them.
- **Insert → Endnote (editor UI).** A new ribbon button under Insert → References
  (next to Insert → Footnote) inserts an endnote at the caret: it places a
  superscript reference run carrying the existing `endnoteRef` field and creates an
  empty note body in `Document.endnotes`, dropping the caret into that body for
  immediate typing — mirroring the footnote command. Later endnote references are
  renumbered in the same transaction so the marker text stays in document order, and
  the whole insertion is a single undo step. Endnote bodies are now editable like
  footnote bodies (the content ops, paragraph split/merge, and the `setEndnote` op
  locate paragraphs inside `Document.endnotes`). The endnote model and `.docx`
  round-trip already shipped; this wires up authoring from the editor only.
- **Table Properties UI — cell/row/table-level formatting (editor).** The existing
  Table Properties dialog (right-click a table cell → *Borders & Shading…*) now exposes
  the table model fields that previously round-tripped but had no editor control: cell
  **vertical alignment** (`w:vAlign` top/center/bottom) and **text direction**
  (`w:textDirection`), **row height** (`w:trHeight` at-least/exact) plus **keep-together**
  (`w:cantSplit`) and **repeat-as-header-row** (`w:tblHeader`) toggles, table **indent**
  (`w:tblInd`), and the table-level **default borders / shading / cell margins**
  (`w:tblBorders` / `w:shd` / `w:tblCellMar`). Quick **context-menu** items cover cell
  alignment and the row toggles. Each edit applies live over the selected cells/rows with
  its own undo step, backed by new selection commands (`setCellVAlignCmd`,
  `setCellTextDirectionCmd`, `setRowHeightAtSelectionCmd`, `setRowPropsCmd`,
  `setTablePropsAtSelectionCmd`) that reuse the `setTableStructure` / `setRowHeight` ops
  and one new table-level `setTableProps` op. Editor-only — the model, import/export, and
  C# bindings already shipped, so `.docx` round-trips are unchanged.
- **Paragraph dialog (editor UI) — borders & shading, line-spacing rule, contextual
  spacing & paragraph flags.** The previously-disabled **Paragraph shading** and
  **Paragraph borders** ribbon buttons (Home ▸ Paragraph) and a new **Line Spacing
  Options…** entry in the line-spacing menu now open a draggable Paragraph dialog
  (`ui/paragraphDialog.ts`) wired to the existing `ParaStyle` fields over the selection
  (one OK = one undo step, via `setParaProps`). It covers paragraph **borders**
  (per-edge top/right/bottom/left/between with a color/width/style spec, reusing the
  table border-picker pattern) + **shading** fill, the line-spacing **rule** (multiple
  / at-least / exactly → `lineRule` + `lineHeightPx`), **contextual spacing**,
  **widow/orphan control**, **vertical text alignment** (`textAlignment`
  top/center/bottom/baseline), **mirror indents**, **suppress line numbers**, and
  **adjust right indent**. Controls are seeded from the caret paragraph via the new
  `editor.currentParaStyle()`. Editor-UI only — the model fields and `.docx`
  round-trip already shipped; no model/import/export or C# change.
- **Interactive image crop (editor UI).** A selected image's context menu gains a
  **Crop** entry that enters crop mode: the full source is shown dimmed behind a
  bright, draggable window with 8 handles (mirroring the image-resize handles). The
  whole session previews purely in the DOM overlay — `Esc`, a click away, or
  selecting another object commits the new insets as a single undoable op (the same
  one-step protocol as drag-to-resize / drag-to-resize-row-height). A **Reset Crop**
  entry (shown only when the image is cropped) clears it. Crop writes the existing
  `ImageBlock.crop` field (OOXML `a:srcRect`, round-trip shipped in #63) via a new
  `setImageCropCmd` reusing the `setImageProps` op — no model/import/export change.
- **Minor run typography & effects (`CharStyle` w:rPr extras).** A grouped set of
  lower-frequency run properties now round-trip and (where visual) paint: double
  strikethrough (`w:dstrike` → `doubleStrikethrough`, two rules), baseline
  raise/lower (`w:position` → `positionPx`, distinct from sub/superscript — it shifts
  without shrinking the font), character width scaling (`w:w` → `widthScalePct`, which
  horizontally stretches/condenses glyphs and is reserved in layout so neighbours don't
  overlap), a kerning threshold (`w:kern` → `kerningMinPx`), emphasis marks (`w:em` →
  `emphasisMark`), the `outline`/`shadow`/`emboss`/`imprint` text effects, a run border
  (`w:bdr` → `runBorder`, reusing the `CellBorder` value type), and a fitText width
  (`w:fitText` → `fitTextPx`). Import parses each, export re-emits them, and `styleEq`
  compares them so styled runs never merge with plain ones. Double strike, position, and
  width scaling paint in both the canvas renderer and PDF export; the remaining effects
  degrade gracefully (preserved, painted as normal text) per the OOXML grouping.
  Authorable from the builder (`paragraph(...).effects({ … })`) and the C# bindings
  (`ParagraphBuilder.Effects(RunEffectsOptions)` + the `EmphasisMark` enum), and
  demonstrated in the default sample document and the C# showcase. Runs without any of
  these fields serialize and paint exactly as before — no drift.
- **Minor paragraph properties (`ParaStyle.widowControl` / `suppressLineNumbers` /
  `textAlignment` / `mirrorIndents` / `adjustRightInd`).** Five lower-frequency `w:pPr`
  children now parse, export, and round-trip through `.docx` where before they were
  silently dropped: `w:widowControl` (widow/orphan control — Word's default ON, honored
  by the pagination engine so an explicit `w:val="0"` lets a lone first/last line break
  across a page boundary), `w:suppressLineNumbers`, `w:textAlignment` (vertical alignment
  of the glyphs within each line box — `top`/`center`/`bottom` hug the respective edge of
  a tall line while `baseline` rides the shared baseline, honored by the layout engine),
  `w:mirrorIndents`, and `w:adjustRightInd`. Authorable via the builder
  (`paragraph(...).widowControl(false).textAlignment("bottom").suppressLineNumbers()
  .mirrorIndents().adjustRightInd()`) and the C# bindings (`ParagraphBuilder.WidowControl`,
  `.SuppressLineNumbers`, `.TextAlignment(LineVAlign)`, `.MirrorIndents`, `.AdjustRightInd`),
  and demonstrated in the default showcase document.
- **Miscellaneous OOXML round-trip — symbols, image crop, in-cell floating images & default
  tab stops (#63).** A grouped backlog of previously-dropped Word features now round-trips:
  - **Symbol characters (`CharStyle.symbol`, OOXML `w:sym`).** Inline symbol-font runs carry
    their font + hex code point (e.g. Wingdings `F0E0`) instead of being dropped or flattened
    to a stray Private-Use character. The run's text is the decoded glyph (painted in the symbol
    font); export re-emits `w:sym`. Authorable via `paragraph(...).symbol(font, charHex)` and the
    C# `ParagraphBuilder.Symbol(font, charHex)`.
  - **Image cropping (`ImageBlock.crop`, OOXML `a:srcRect`).** DrawingML crop insets (1/1000 of a
    percent) are parsed into 0..1 fractions, painted (canvas source-rect crop; PDF clip + scaled
    draw), and re-emitted as `a:srcRect`. Authorable via the `image()` `crop` option and the C#
    `ImageOptions.Crop` / `ImageCrop` record.
  - **Floating images inside table cells.** Anchored (`wp:anchor`/`wrapNone`) images inside a
    `w:tc` are preserved as cell `ImageBlock`s instead of being dropped (cells previously kept
    only paragraphs/tables).
  - **`settings.xml` (`Document.defaultTabStopPx` + `Document.compatSettings`).** `w:defaultTabStop`
    is honored at layout — a `\t` past the last explicit tab stop now advances by the document's
    interval instead of a fixed 0.5in constant — and `w:compat/w:compatSetting` triples round-trip
    verbatim. Authorable via `DocumentBuilder.defaultTabStop(px)` / C# `DefaultTabStop(px)`.
  All four are demonstrated in the default showcase document and the C# showcase, with docx
  round-trip + layout tests.
- **Drag-to-resize table row height.** A horizontal grip on each table row's bottom
  edge can now be dragged to set that row's height, mirroring the existing
  column-width resize interaction. The pointer shows a `row-resize` cursor and an
  accent guide over a grabbable boundary; the drag previews live (each frame
  relayouts the table) and commits one undoable step on drop, writing the dragged
  pixel height into `TableRow.height` as `{ value, rule: "atLeast" }` by default (an
  existing `exact` rule is preserved). Layout still floors an `atLeast` row at its
  content height, so a row never drags below its content. Column grips win at a cell
  corner, so the two interactions never collide. Editor-UX only — it reuses the
  `TableRow.height` model (and thus the `.docx` round-trip) added previously; built
  on a new `setRowHeight` op with an undo inverse.
- **Minor & advanced table properties (issue #61).** Tables and cells now carry a set of
  previously-dropped OOXML properties. Table-level: `TableBlock.indentPx` (`w:tblPr/w:tblInd`)
  indents the whole table from the leading margin, and `TableBlock.bidiVisual`
  (`w:tblPr/w:bidiVisual`) lays the columns out in **right-to-left** visual order (grid column
  0 paints at the right edge) — both honored by the layout engine; `TableBlock.overlap`
  (`w:tblOverlap`) plus `caption`/`description` (`w:tblCaption`/`w:tblDescription` alt text)
  round-trip as metadata. Cell-level: `TableCell.textDirection` (`w:textDirection` — vertical
  text), `noWrap` (`w:noWrap`), `fitText` (`w:tcFitText`), and `hideMark` (`w:hideMark`) all
  round-trip through `.docx` (parsed in `documentParser`, emitted from `documentXml` in
  `CT_TblPrBase`/`CT_TcPr` child order). Authorable via the builder
  (`.table(rows, { indent, bidiVisual, overlap, caption, description })` and cell
  `{ textDirection, noWrap, fitText, hideMark }`) and the mirrored C# bindings
  (`TableOptions` + `CellOptions`/`CellSpec`, new `CellTextDirection`/`TableOverlap` enums),
  and demonstrated in the default showcase document. (Floating positioned tables `w:tblpPr`
  and conditional-band cell margins `w:tblStylePr/w:tcMar` remain out of scope for this pass.)
- **Endnotes (`Document.endnotes` + `CharStyle.endnoteRef`).** Endnotes now round-trip
  alongside footnotes: `endnotes.xml` and `w:endnoteReference` markers import into a
  per-document note store (previously dropped on import), export back out with their own
  content-type override, relationship, and part, and lay out at the **end of the document**
  under a separator rule (Word's "end of document" placement — the counterpart to the
  page-bottom footnote area). Reference markers auto-number in document order, render in the
  canvas and PDF, and notes collect on continuation pages as needed. Authorable via the builder
  (`paragraph(...).endnote("…")` / callback form) and the C# bindings (`ParagraphBuilder.Endnote(…)`),
  and demonstrated in the default showcase document.
- **Section break parity (`evenPage`/`oddPage`) + line numbering (`w:lnNumType`).** Section
  breaks now carry an OOXML `w:sectPr/w:type` of `nextPage` (the default), `evenPage`, or
  `oddPage` (`Paragraph.style.sectionBreak.type`): an even/odd break forces the new section's
  first page onto an even/odd page number, with the layout inserting a blank filler page when
  the running page count has the wrong parity (Word's behavior). Sections can also enable line
  numbering (`SectionProps.lineNumbering` / `SectionPatch.lineNumbering`, OOXML `w:lnNumType` —
  `countBy`/`start`/`restart` `continuous`/`newPage`/`newSection`/`distance`), printing a number
  in the margin beside each body line; previously the break type was silently flattened to
  `nextPage` and `w:lnNumType` was dropped on import. Both parse, round-trip through `.docx`,
  and render in the canvas renderer and PDF export. Authorable via the builder
  (`sectionBreak({ breakType, lineNumbering })`, `pageSetup({ lineNumbering })`) and the C#
  bindings (`SectionBreakOptions.BreakType`/`LineNumbering`, `PageSetup.LineNumbering`), and
  demonstrated in the default showcase document.
- **Paragraph borders & shading (`ParaStyle.borders` + `ParaStyle.shading`).** A whole
  paragraph can now carry a border box (OOXML `w:pBdr` — `top`/`right`/`bottom`/`left`,
  plus a round-tripped inter-paragraph `between` edge) and a background fill (paragraph-level
  `w:shd`), where previously only table **cell** shading/borders were supported. Each border
  edge reuses the table `CellBorder` value type (color + width + `single`/`double`/`dashed`/
  `dotted`). Both round-trip through `.docx` (`w:pPr/w:pBdr` + `w:pPr/w:shd`) and render
  pixel-exact in the canvas renderer and PDF export — the fill paints beneath the text (under
  selection/search highlights, like cell fills) and the box hugs the paragraph between its
  indents; the border line widths are reserved in the surrounding block gaps. It renders in
  every paragraph context (body, floats, table cells, headers/footers, footnotes). Authorable
  via the builder (`paragraph(...).borders({...}).shading("#rrggbb")`) and the C# bindings
  (`ParagraphBuilder.Borders(ParaBorders.All(...))` / `.Shading(...)`), and demonstrated in
  the default showcase document.
- **Paragraph contextual spacing (`ParaStyle.contextualSpacing`).** Paragraphs flagged
  with OOXML `w:contextualSpacing` now drop their before/after spacing against an
  adjacent paragraph of the **same style** — Word's default for list styles, so
  same-style runs (list items, verse stanzas) sit tight while the run's outer edges
  keep their spacing. Previously the flag was ignored, so imported lists over-spaced.
  Parsed in `props.ts`, baked through the style cascade onto the concrete paragraph,
  re-emitted on export both as direct paragraph formatting (`paraCoreXml`) and as a
  paragraph-style delta (`partialPPrXml`); the layout engine suppresses the spacing
  between adjacent same-`namedStyle` paragraphs. Authorable via the builder
  (`ParagraphBuilder.contextualSpacing()`) and the mirrored C# bindings
  (`ParagraphBuilder.ContextualSpacing` / `ParaStylePatch.ContextualSpacing`);
  demonstrated in the default sample document and the C# showcase.
- **Table row properties (`TableRow.props` — `w:trPr`).** Rows now carry their own
  OOXML row properties, previously dropped entirely (`TableRow` was just `{ cells }`):
  `height` (`w:trHeight`) pins a fixed/minimum row height — `rule: "atLeast"` grows the
  row with its content, `rule: "exact"` forces the height (taller content clips);
  `cantSplit` (`w:cantSplit`) keeps a row whole across a page break — the paginator is
  row-atomic (it only ever breaks a table between rows, never inside one), so this
  invariant already holds for every row and the flag round-trips for `.docx` fidelity; and
  `repeatHeader` (`w:tblHeader`) re-draws the leading contiguous header rows at the top
  of every page a table continues onto. All three round-trip through `.docx` (parsed in
  `documentParser`, emitted from `documentXml` as `w:trPr`) and drive the layout engine
  (`measureTable` honors the height; `placeTableChunked` repeats the header band as its
  own contiguous placed block per continuation page). Authorable via the builder
  (`TableBuilder.row(cells, { height, heightRule, cantSplit, header })`) and the mirrored
  C# bindings (`RowOptions` + `RowHeightRule`); demonstrated in the default sample
  document and the C# showcase.
- **Run case transforms (`CharStyle.caps` + `CharStyle.smallCaps`).** All-caps
  (OOXML `w:caps`) and small-capitals (`w:smallCaps`) now parse, model, render and
  round-trip — common on headings and styles. The model text is untouched; the
  layout bakes the transform into throwaway display runs (uppercasing every letter,
  and for small caps splitting the originally-lowercase letters into reduced-size
  sub-runs), so the canvas renderer and the PDF painter draw the transformed glyphs
  with no painter-specific code. The transform is offset-transparent — it preserves
  the UTF-16 length per code point — so the caret, hit-testing and measurement land
  on the uppercased glyphs. Parsed in `props.ts` (a `w:rPr` toggle, including the
  style-cascade XOR), emitted in `styleProps.ts`, transformed in the layout prepare
  cache. Authorable via the builder (`.caps()` / `.smallCaps()`) and the mirrored C#
  bindings (`Caps` / `SmallCaps`); demonstrated in the default sample document and
- **Fixed line spacing (`ParaStyle.lineRule` + `lineHeightPx`).** Line spacing can now
  be a fixed point height, not just a multiplier of the font size. OOXML
  `w:spacing/@w:lineRule="exact"` pins every line box to `lineHeightPx` (taller glyphs
  clip); `"atLeast"` floors the height there but lets a taller line grow. Previously
  both were dropped on import (with a `line-rule-exact` warning) and re-exported as
  `auto`. Now parsed in `import/docx/props.ts` (the `w:line` value read as twips, not
  240ths), honored in the layout engine's line-metrics (so it drives pagination), and
  re-emitted with the correct `w:lineRule` on export — a full round-trip. Authorable via
  the builder (`SpacingOptions.lineRule` / `.lineHeightPx`) and the mirrored C# bindings
  (`LineRule` enum on `SpacingOptions`); demonstrated in the default sample document and
  the C# showcase.
- **Run-level round-trip fidelity — character tracking, theme tint/shade, and CS/EA
  font slots.** Three details Word stores on a run now survive a full `.docx`
  round-trip. (1) **Character tracking** (`w:spacing`): the importer reads run-level
  `w:spacing` into `CharStyle.letterSpacingPx` — export already emitted it, so authored
  tracking was lost inbound until now. (2) **Theme tint/shade** (`w:themeTint` /
  `w:themeShade`): a theme color with a tint/shade now resolves to its actual lighter
  /darker shade instead of flattening to the flat base hue (applied per RGB channel in
  `theme.ts`). (3) **Complex-script & East-Asian font slots** (`w:rFonts/@w:cs` and
  `@w:eastAsia`, plus their themed variants): captured into the new
  `CharStyle.fontFamilyComplexScript` / `fontFamilyEastAsia` (only when they name a face
  distinct from the Latin slot, since Word writes `w:cs = w:ascii` for plain Latin runs),
  emitted back from `documentXml`. Authorable via the builder (`fontComplexScript()` /
  `fontEastAsia()`, plus the existing `letterSpacing()`) and the mirrored C# bindings
  (`ParagraphBuilder.FontComplexScript` / `FontEastAsia`, `CharStyle.FontFamilyComplexScript`
  / `FontFamilyEastAsia`); demonstrated in the default sample document and the C# showcase.
- **Table cell vertical alignment (`TableCell.vAlign`).** Cells can now align their
  content to the `top` (default), `center`, or `bottom` of the cell box via OOXML
  `w:tcPr/w:vAlign` — previously content always hugged the top regardless of the
  source `w:vAlign`. The layout engine offsets the cell's block stack by the slack
  between the content height and the (often taller) painted cell box, so it is most
  visible in a tall `rowSpan` cell or beside a tall sibling row. Round-trips through
  `.docx` (parsed in `documentParser`, emitted from `documentXml` after `w:tcMar` per
  `CT_TcPr`; an explicit `top` normalizes to absent). Authorable via the builder
  (`CellSpec.vAlign` / `CellOptions.vAlign`) and the mirrored C# bindings
  (`CellVAlign` enum); demonstrated in the default sample document and the C# showcase.
- **Underline style & color (`CharStyle.underlineStyle` + `CharStyle.underlineColor`).**
  Underlines now carry their OOXML line style (`w:u/@w:val` — `double`, `thick`, `dotted`,
  `dash`, `dotDash`, `dotDotDash`, `wave`) and an optional color (`w:u/@w:color`, incl.
  theme colors resolved at import). Previously `w:u` round-tripped as a boolean and always
  re-exported as `w:val="single"`, painting only a solid line. Import parses the real
  `w:val` (folding Word's heavy/long variants onto the nearest base style) and resolves a
  themed underline color through `theme1.xml`; export emits the true `w:val` + color; both
  the canvas renderer and PDF painter draw double/dotted/dashed/dot-dash/thick/wave rules
  and honor the underline color (hyperlink affordances still paint a plain rule in the link
  color). Authorable from the builder (`.underline(true, { style, color })`) and the C#
  bindings (`Underline(on, style, color)` / `CharStyle.UnderlineStyle`/`UnderlineColor`),
  demonstrated in the default showcase. Runs without a style/color serialize and paint
  exactly as before — no drift.
- **Table-level default borders / shading / cell-margins round-trip (`TableBlock.defaultBorders`
  / `defaultShading` / `defaultCellMargin`).** A table's `w:tblPr/w:tblBorders` (including the
  interior `insideH`/`insideV` edges), `w:tblPr/w:shd` default fill and `w:tblPr/w:tblCellMar`
  default cell padding were parsed on import and cascaded onto cells, but **never re-emitted at
  the `tblPr` level** — only the resolved per-cell `w:tcBorders`/`w:shd`/`w:tcMar`. A
  Word→edit→Word cycle therefore dropped the table-wide defaults. The model now retains them and
  export hoists them back into `w:tblPr` (respecting `CT_TblPrBase` child ordering:
  `tblBorders → shd → tblLayout → tblCellMar`), so they survive the round-trip while per-cell
  overrides still win. Authorable via the builder (`.table(rows, { borders, shading, cellMargin })`,
  mirrored in the C# `TableOptions` + `TableBorders`), shown in the default showcase document, and
  covered by an import→export→re-import test. Tables without these fields serialize byte-identically.
- **Table width & alignment (`TableBlock.preferredWidth` + `TableBlock.align`).** Fixed
  tables can now be narrower than the page and positioned within it, instead of always
  spanning the full content width (the only previous escape was AutoFit to Contents).
  `preferredWidth` sizes the whole grid — `{type:"pct"}` as a percentage of the content
  width or `{type:"px"}` as an absolute width, clamped to `[ncols × 24px, contentWidth]`
  and applied in fixed layout only; the columns keep their proportions. `align`
  (`left`/`center`/`right`) shifts the table within its band **whenever it is narrower
  than the page**, so it also aligns AutoFit-to-Contents tables. Both round-trip through
  `.docx` (`w:tblPr/w:tblW` as `dxa`/`pct` + `w:tblPr/w:jc`) and render pixel-exact in
  PDF export (shared layout engine). Reachable from three surfaces: the right-click
  **AutoFit & Size** submenu (quick 25/50/75/Full presets + alignment), the standalone
  ribbon's AutoFit dropdown, and the **Table Properties → Table size** section (free-form
  value in % or inches + alignment). New ops `setTablePreferredWidth` / `setTableAlign`;
  commands `setTablePreferredWidthAtSelectionCmd` (also pins the table to fixed layout)
  and `setTableAlignAtSelectionCmd`. Tables without these fields serialize and lay out
  byte-identically — no export drift.
- **CJK (Chinese) text in export.** A subset of **Noto Sans SC** (OFL; ~3,755 common
  GB2312 Level-1 hanzi + ASCII + CJK punctuation, ~1.4 MB) now ships as a built-in
  fallback face (`NotoSansSC`). CJK runs are script-split onto it automatically, so
  Chinese text no longer renders as `.notdef`/tofu (the "x" boxes) in PDF export — it
  works with **zero configuration** across the editor, the browser export worker, and
  the headless backend (`/render.pdf`). The face is a single Regular (every style maps
  to it), registered like the bundled math font; the editor loads it **lazily** and
  re-lays-out when it arrives (new `Editor.refreshFonts()`), so Latin-only documents
  pay no first-paint cost. `CjkConfig.fallbackFont` now defaults to it; set a custom
  font's family to override, or `""` to opt out (keeps the browser's on-screen system
  fallback, CJK then tofu in PDF). DOCX is unchanged — it writes the document's own
  family and Word substitutes its CJK font. Builds on the existing CJK line-breaking
  and `scriptSplitRuns` fallback machinery.
- **Mathematical equations (MathML).** First-class math, MathML-native end to end:
  - **Display & inline equations.** Block equations (`EquationBlock`) sit on their own
    line (left/center/right aligned); inline equations ride inside a text line via a
    `U+FFFC` sentinel run carrying `CharStyle.equation`, so every caret / selection /
    offset invariant is preserved.
  - **STIX Two Math typesetting.** The bundled STIX Two Math font (OFL), driven by the
    font's real OpenType **MATH** table constants (extracted at build time), typesets
    fractions, radicals, sub/superscripts, n-ary operators, matrices, accents, and
    growing delimiters. Identifiers remap to the Mathematical Alphanumeric block so
    italic / bold / blackboard / script glyphs are true glyphs (no faux slant).
  - **Visual editor (MathML or LaTeX).** A floating equation editor with a template /
    symbol palette and a live canvas preview; accepts Presentation MathML or a common
    LaTeX subset, switchable both ways. Insert via the ribbon (Insert → Equation) or
    right-click → **"Edit Equation…"** on any existing equation; Display/Inline toggle
    and alignment controls.
  - **`.docx` round-trip (OMML).** Equations import from and export to OMML — display
    `m:oMathPara` and inline `m:oMath`, including under/over limits, `m:scr` script
    variants (double-struck / fraktur / script / sans-serif / monospace), `mspace`
    widths, and hyperlinked inline equations. PDF export paints equations vectorially
    (with run-style decorations and links).
  - The default showcase document gains a **"Mathematics — MathML equations"** section.

### Fixed
- **Nested bullet glyphs no longer render as tofu in PDF export.** The default
  bullet levels are `•`/`◦`/`▪` (`shared/src/model/lists.ts`), but `◦` (U+25E6
  WHITE BULLET) and `▪` (U+25AA BLACK SMALL SQUARE) are absent from the bundled
  PDF Latin font subset, so the PDF painter drew `.notdef`/tofu for the nested
  levels (the on-screen canvas was fine because the browser's system font has the
  glyphs). The three standard bullets now paint as **vector shapes** (filled disc /
  hollow ring / filled square) in BOTH the canvas renderer and the PDF painter via a
  shared `bulletShapeFor` geometry, so they no longer depend on font glyph coverage
  and look identical on screen and in the export. Numbered and custom-glyph markers
  still paint as text. (#99)
- **Paragraph border (`w:pBdr`) no longer overlaps the text.** A paragraph border box
  was stroked exactly at the text content edges with no border-to-text padding, so a
  wide or `double` rule straddled the glyphs (visible with the showcase's left-only
  double orange border). The decor box is now expanded **outward** by a small
  border-to-text padding (default ~4px) on every side in **both** painters
  (`paint/renderer.ts` canvas + `export/pdf/paintBlock.ts` PDF), derived from a single
  shared `paraDecorBox` helper so canvas and PDF stay pixel-identical. The padding is
  reserved in layout (`paraDecorFor` + the body-flow block-gap reservation) so adjacent
  content can't collide with the box, and the shading fill (`w:shd`) now matches the
  padded box so it reaches the border. Shading-only paragraphs keep their text-edge box
  (no padding) and render byte-identically. No model/import/export change.
- **PAGE / NUMPAGES field inside a table now updates when the table moves pages.**
  Moving a table to a new page (e.g. a `Ctrl+Enter` page break before it) left a
  body PAGE field in one of its cells showing the old page number until an unrelated
  edit — typically a column resize — happened to bust the table's measurement cache.
  The per-page token-resolution pass rewrote the `{page}` fragment's text **in place**
  on `LineBox` objects that the line/table caches alias, baking the resolved number
  into the cache; a later layout that reused the cache verbatim (table revision + width
  unchanged) then re-painted the stale number, since the cached fragment already read
  `"1"` and no longer matched the `{page}` token. Resolution is now **clone-on-write**:
  any line carrying a page field is copied before substitution, so the cached `{page}`
  token stays pristine and the field re-resolves correctly on every pass. The same
  latent staleness affected body paragraphs whose PAGE field reflowed across a page.
- **Nested display equations (table cells / header-footer bands) are now editable.**
  The `setEquation` / `setEquationAlign` ops and the editor lookups (`setEquationAlignCmd`,
  the right-click menu, object-selection delete) only scanned top-level `doc.blocks`, so
  a display `EquationBlock` imported into a table cell or a header/footer band — e.g. from
  a `.docx` `m:oMathPara` in a cell — couldn't be edited / aligned / deleted (the op threw
  "not found"). They now resolve equations through a container-aware locator
  (`locateEquation`/`locateBlock`, the generalization of `locateImage`) that searches the
  body, every band story, and table cells in either; layout already rendered them there.
- **Enter inside a block-level content control.** Pressing Enter in a paragraph or
  table cell wrapped by a block-level SDT (e.g. a "section" control) was swallowed —
  the guard treated every control at the caret as inline. Now only *inline* controls
  block the split; block-level controls hold multiple paragraphs (like Word), and the
  new paragraph inherits the control (`splitParagraph` carries the block `sdtPath` to
  the tail, with the merge inverse restoring it so undo round-trips).
- **Caret / selection on multi-run RTL paragraphs.** Lines store their fragments in
  visual order, so on a bidi-reordered line with more than one run the flattened line
  index took its start/end offset from the visually-, not logically-, first/last
  fragment. The caret jumped to the previous line and whole-paragraph selection broke
  on RTL paragraphs mixing runs (e.g. Arabic text with embedded `PAGE`/`DATE` fields).
  The line index now derives each line's logical span from the min/max offset across
  its fragments. (Single-run and LTR lines are unaffected.)
- **Tab stops in RTL paragraphs.** Tab-stopped lines were always laid out left-to-right
  from the left margin, so in an RTL (`w:bidi`) paragraph the stops measured from the
  wrong edge and the cells ran the wrong way (the tabs-on-an-RTL-line case was left in
  logical order). An RTL tab line is now reflected about the content box: stop positions
  are measured from the **right (start) edge**, tab cells fill **right-to-left**, and
  right/center/decimal alignment plus leaders and tab arrows mirror with them. Each cell
  keeps its own embedding level, so RTL-script cells right-anchor and caret / hit-testing
  take the bidi path. The LTR fast path is byte-identical. (Issue #6.)

## [0.7.5] — 2026-06-25

### Added
- **CJK & bidirectional (RTL) text.** The layout engine now lays out East-Asian and
  right-to-left scripts the way Word does — all measured on canvas:
  - **CJK line-breaking & kinsoku** work out of the box (between Han/Kana characters
    with no spaces, no line-start/-end punctuation), via the pretext analyzer.
    `WordCanvas({ cjk: { locale: "ja" | "ko" | "zh" } })` tunes locale-specific
    breaking; `cjk.fallbackFont` (a registered `fonts` family) routes CJK runs to a
    known font for consistent measurement, PDF/DOCX export, and embedding.
  - **Bidirectional layout (UAX #9).** New `ParaStyle.direction` (`"ltr"`/`"rtl"`,
    OOXML `w:bidi`) and `CharStyle.rtl` (`w:rtl`). RTL paragraphs reorder runs into
    visual order, right-align by default, and mirror left/right indents; caret,
    hit-testing, selection rectangles, and Left/Right arrow keys all follow the
    visual order. List markers hang in the mirrored right-side gutter, justified
    RTL lines fill edge-to-edge, and left/right indents mirror. A **LTR/RTL**
    toggle was added to the ribbon's Paragraph group, and `w:bidi`/`w:rtl`/`w:jc`
    round-trip through `.docx`.
  - The default showcase document gains an **"International text — CJK &
    bidirectional"** section demonstrating all of the above.
  - *Implementation note:* the bidi reorder reuses pretext's Unicode Bidi
    Algorithm via two small, documented patches under `/pretext-patch`
    (applied by `patch-package` on `postinstall`).
- **Inspector: ergonomics.** Keyboard navigation in the trees (↑/↓ move, →/←
  expand/collapse, Enter reveals), a right-click context menu (copy block id / JSON
  / label, reveal), a **Follow caret** toggle that scrolls the tree to the node
  under the caret as you edit, and toolbar buttons to copy the **Document** /
  **LayoutTree** JSON or **load** a Document JSON snapshot.
- **Inspector: History tab (edit log).** A new tab streams the recorded change log
  — each entry's origin (typing/command/paste/undo/redo), op summary, and time —
  click to inspect its ops; Undo / Redo buttons step the document from the panel.
- **Inspector: Probe tab (hit-test readout).** A new tab shows, live under the
  pointer, exactly what the input layer resolves there — the page point, caret
  position (block id + offset), content-control chain, field, and table cell — with
  a Freeze chip to pin a reading. New `Editor.setInspectorProbe(active)` +
  `onInspectorProbe` callback + `InspectorProbe` type.
- **Inspector: Layout tab (geometry tree).** A new tab renders the laid-out geometry
  — pages → placed blocks → lines → fragments, and tables → rows → cells → blocks —
  each node showing its position/size and highlighting its exact painted rect on
  hover (fragment-precise). New `{ kind: "rect" }` inspector target.
- **Inspector: Problems tab (model validator).** A new tab runs a pure model walk
  and lists integrity problems grouped by severity — dangling content-control /
  field / style / list / table-style / bookmark / TOC / footnote references,
  duplicate block ids, and unused content controls / fields — each linking to the
  offending block. The tab shows an error+warning count badge.
- **Inspector: canvas layout overlays.** The develop-mode panel gains a row of
  toggle chips that draw the layout structure directly on the page — block boxes,
  line boxes, inline-fragment boxes, baselines, table cell boxes, the content
  margins, and a per-page info badge — for debugging "why does this render like
  that." Drawn from the live layout tree on every repaint (survives scroll/zoom),
  cleared when the panel closes. New `Editor.setDebugOverlay(kind, on)` + a paint
  channel. The inspector panel was also refactored into a tabbed shell (Model tab
  today; Layout / Problems / History to follow).
- **Wrap a selected image in a content control (picture content control authoring).**
  The *Rich text content control* ribbon button and a new *Wrap in Content Control*
  image right-click entry now act on a selected image, tagging the image block with a
  fresh control id (nesting-aware — it appends to any existing ancestry) and
  registering the control. Previously content controls could only be authored around
  a text selection, so picture content controls could only arrive via import. The
  control's frame + breadcrumb appear immediately, and *Content control properties*
  and *Remove content control* work on the image's control too (the ribbon exposes the
  active control via `Editor.activeContentControlId()`). `removeContentControl` now
  strips block-level membership off image blocks (top-level and in table cells), so
  wrap/unwrap round-trips cleanly. Works for an image anywhere it's selectable —
  body or table cell.
- **Develop mode via the `develop` constructor option (`develop: true`).** Reveals a
  dedicated **Developer** ribbon tab whose *Inspect document tree* button opens a
  floating, draggable, devtools-Elements-style panel over the parsed `Document`
  model. The tree covers body blocks → runs, tables → rows → cells, header/footer
  bands, footnotes, and the side-tables (styles, lists, table styles, content
  controls, fields, bookmarks). **Content-control and field membership is
  reconstructed as tree structure** — body/cell-level SDTs show as `SDT ->
  paragraph -> runs`, and inline fields as `Field -> run` — instead of flat
  paragraphs/runs. Hovering a node paints a highlight box over its region on the
  canvas (paragraphs, runs, images, **table rows and cells**, content controls, and
  fields all resolve to their painted rects); hovering the page reveals the matching node
  in the tree (reverse sync); clicking selects + scrolls to it (cells and images
  included); selecting shows the node's properties + raw JSON. A filter box narrows
  the tree. It's a pure debugging aid,
  gated twice over: the tab only exists when the flag is set, and nothing dev-related
  runs until the panel is opened from it (the canvas↔tree hover signal is dormant
  otherwise). The editor surface gains `setInspectorHighlight(blockId | null)` and
  `setInspectorActive(active)` with an `onInspectorHover` callback, and the paint
  layer a `setInspectorRects` overlay channel. The panel is draggable by its header
  and **resizable** (drag the bottom-right corner). Off by default; leave it off for
  production embeds. The dev harness (`npm run dev` / `dev:online`) and the
  **offline embed example** opt in with `?devMode=true` in the URL; the constructor
  builder example exposes a `develop` toggle.
- **Content-driven autofit tables (`TableBlock.widthMode`).** Tables can now size
  their columns from cell content instead of only the fixed
  `colFractions × content-width` model. Two modes, exposed under the **Table →
  AutoFit** ribbon menu: *AutoFit to Contents* measures each cell's min (widest
  unbreakable token) and max (natural unwrapped) width, solves the grid with the
  CSS automatic-table-layout algorithm, and lets the table **shrink below the page
  width** to fit; *AutoFit to Window* solves the same way then fills the available
  width. Column-spanning cells distribute their demand across the columns they
  cover; per-cell preferred widths (`w:tcW`) clamp a column's content min/max up to
  the preference. Hovering a column border now **highlights the boundary** with a
  vertical accent guide (over a soft grab-zone band) so the resize affordance is
  discoverable; dragging it **cancels autofit** and pins the table back to fixed
  widths (Word's behavior). The default sample document now
  includes an AutoFit-to-Contents table. The mode round-trips through `.docx`
  (`w:tblLayout` + `w:tblW` + `w:tcW`) and renders pixel-exact in PDF export (which
  replays the layout engine). **Export is hint-based** — the writer emits
  `w:tblLayout="autofit"` and relies on Word re-autofitting from the `w:gridCol`
  snapshot rather than writing the solved widths; import is conservative (a table
  adopts autofit only when it explicitly declares it, so existing fixed-proportional
  imports never shift). See *Known limitations* in `README.md`.
- **Custom fonts via the `fonts` constructor option.** Embedders can now supply their
  own fonts, loaded from URLs at runtime, instead of being limited to the bundled
  metric clones: `fonts: { fonts: [{ family, faces: { regular, bold?, italic?,
  boldItalic? }, sizing: { ascent, descent } }], disableBuiltin?: string[] }`. A
  custom family is stored in the model, listed in the toolbar, and rendered as
  itself (never substituted to a clone) across all three contexts that must agree
  for page-accurate layout — the editor canvas, the client-side export worker, and
  the headless Node backend. The required `sizing` (ascent/descent as fractions of
  em — the same role the bundled clones' baked metrics play) is what keeps the
  editor, the browser export, and a server-side export paginating **identically**.
  `disableBuiltin` hides built-in families from the toolbar by their original name
  (e.g. `"Calibri"`) while keeping them loaded and resolvable, so a loaded `.docx`
  that still references them renders. PDF subset-embeds the actual faces used; DOCX
  writes the custom family name. **Backend:** server-side export honors a document's
  saved font config (persisted at `POST /docs`), and `POST /render.pdf` accepts a
  `fonts` part — the server fetches the faces and **caches them on disk** (keyed by
  URL) so they're reused across renders. TTF/OTF only (WOFF/WOFF2 rejected); a missing
  bold/italic/bold-italic face falls back to the regular face in both the editor and
  the exporters. Custom-font state is **per-instance**: each editor mount and each
  export job owns its own registry (threaded onto the export call, not a shared
  module slot), so multiple `WordCanvas` instances with different fonts — and
  concurrent export jobs — never cross-contaminate. The backend font fetcher is
  **bounded** (http/https only, optional host allowlist, timeout, and size cap) to
  prevent SSRF/abuse, and `POST /docs` validates the `fontsConfig` before storing it.
  See `FONTS.md` for the design.
- **New `custom-fonts` example** (`examples/custom-fonts`, served at
  `/examples/custom-fonts`): a no-build offline embed that self-hosts PT Serif (4
  faces) via the `fonts` option, hides the built-in Calibri from the toolbar, and
  exports a PDF with the custom faces embedded.

### Changed
- **Incremental relayout & repaint (large-document responsiveness).** Editing,
  dragging, and resizing on long documents no longer pay the cost of a full
  re-layout and full-canvas repaint per frame. Relayout now reuses cached
  per-page header/footer layout across passes, memoizes top-level table
  measurement by `(revision, width)`, and `afterMutation` skips redundant work on
  transient drag/composition frames; repaint targets only the pages that actually
  changed instead of every live page. Image resize is previewed in a lightweight
  DOM overlay during the drag and committed to the model once on release. Behavior
  is unchanged — the document lays out and paints identically — but interaction
  stays smooth as documents grow. See `perf/baseline.md` for measured results.

### Fixed
- **Images inside table cells can now be resized and aligned.** Selecting an image
  that lives in a table cell — the usual shape of a content-control report template,
  where a `w:sdt` wraps a table and the picture sits in a cell — drew the selection
  frame and resize handles, but dragging a handle or clicking an alignment button did
  nothing. The image-props command (resize, align, wrap) and the live resize-preview
  both looked only at top-level body blocks, so a cell image was silently inert even
  though the underlying op (and every other image command: delete, layer/z-order,
  move) already resolved images in cells. Both paths now use the cell-aware
  `locateImage`, so in-cell images resize and align like any other.
- **Smooth image-resize dragging (no more lag inside large tables).** The resize
  handle ran a full relayout on every `pointermove`. For a top-level image that's a
  few milliseconds and keeps up, but an image in a big table makes each relayout far
  heavier than the gap between move events, so previews queued up and the image
  trailed the cursor by a fraction of a second. Resize previews are now coalesced to
  one per animation frame (the same per-frame throttle pinch-zoom already used), so
  the image stays glued to the handle regardless of table size; the final committed
  size is unaffected. Table **column-boundary drags** were throttled the same way —
  they ran a transient relayout per `mousemove` and lagged identically on big tables.
- **Selecting an image inside a content control now shows that it's in one.** When a
  control's only content was an image, selecting the image (object selection clears
  the text caret, which is what drove the chrome) left no frame, no breadcrumb, and a
  greyed-out Controls ribbon group — the user had no way to tell the picture was
  inside a `w:sdt`. Object selection now derives the control ancestry from the image
  (its block-level `sdtPath`, plus the wrapping table's for a cell image), so the
  control frame + breadcrumb tab render and the *Content control properties* / *Remove*
  buttons light up. Image-only controls that sit in a table cell — which the
  block-level and run-based frame paths both miss — are framed around the image.

## [0.7.4] — 2026-06-23

### Added
- **Nested content controls (`w:sdt` inside `w:sdt`) now round-trip.** Content
  controls can nest — an outer "section" control wrapping inner field controls,
  inline-in-inline, block-in-block, inline-in-block, and a control wrapping a whole
  table (with its own inner controls per cell). Import reconstructs the full
  ancestry, export re-emits the nested `w:sdt` structure faithfully, and editing
  inside an inner control preserves its outer ancestry. Previously nesting
  collapsed to a single level and a multi-block/section control fragmented into
  several controls on export. Complex Word fields (PAGE/DATE/IF…) nest inside
  controls too — a field result inside nested controls keeps both its control
  ancestry and its field membership across the round-trip. The default showcase
  document now demonstrates this: an outer-wraps-inner inline control, plus a
  block-level "section" control around a paragraph and a table whose value cell
  holds its own inner control. No other browser-native Word editor preserves
  nested content controls — generated C#/OOXML report section→field structure
  survives an editor round-trip intact.
- **Hover highlighting for content controls (including nested).** Pointing at a
  control now frames it without moving the caret; nested controls draw as
  concentric frames (outer→inner, depth-graded) with a single breadcrumb tab
  (e.g. "Section › Appraisal Fee"), so the hierarchy is visible at a glance. The
  caret still frames the control it sits in; hover takes precedence while pointing.
  Block-level controls (incl. ones wrapping a whole table) frame their full span.
- **Responsive compact ribbon + scrollable tab strip.** The ribbon now adapts to a
  narrow editor: a `ResizeObserver` on the editor root toggles a dense, single
  horizontally-scrollable group row with hidden group labels below ~720px wide
  (720 enter / 760 exit hysteresis). It keys off the editor's **own** width — so it
  also compacts when embedded in a narrow pane on a wide page, not just on small
  devices. The tab strip (`Home`/`Insert`/…) is now horizontally scrollable on
  overflow (swipe, trackpad, or mouse-wheel-to-horizontal) while the mode-select /
  Review / collapse-chevron cluster stays pinned at the right.
- **Desktop app — `desktop/` Tauri v2 workspace.** WordCanvas can now ship as a
  native desktop application. The editor runs fully offline (no backend calls);
  all actions live in the ribbon — **Open `.docx`** uses the native OS file picker
  and **Export DOCX/PDF** routes through a native Save dialog (the `onSave` hook),
  so there's no separate native menu. It builds to a single **portable `.exe`**
  with no installer (`bundle.active: false`, statically-linked CRT);
  `.github/workflows/desktop-release.yml` produces it on `windows-latest`. The
  desktop frontend consumes the published `@forevka/wordcanvas` bundle as an
  external dependency (the `examples/embed-live` pattern), so the package itself
  is unchanged. See `desktop/README.md`.

### Changed
- **Content-control membership moved from a scalar `CharStyle.sdtId` to an ordered
  `CharStyle.sdtPath` (plus block-level `Block.sdtPath`).** This is what enables
  nesting: a run's controls are an outer→inner path, and a run's full enclosing
  chain is `block.sdtPath ++ run.style.sdtPath`. **Behavior change, no migration
  shim:** documents loaded from a `.docx` are unaffected (import populates
  `sdtPath` directly). Old **persisted snapshots / op logs** written before this
  change carried the legacy scalar `sdtId`, which is no longer read — their
  content-control *membership* is dropped on load (the text is intact); re-import
  from the source `.docx`, or re-save, to restore controls as `sdtPath`. Code
  reading membership must use `sdtPath` / the `@cw/shared` sdt helpers
  (`innermostSdtId`, `fullSdtChain`, …) instead of `style.sdtId`.
- **Export shows a busy overlay while rendering.** The DOCX/PDF Export buttons now
  display an "Exporting…" overlay during the (potentially multi-second) render —
  matching the existing open-document overlay — so the editor no longer appears
  frozen mid-export.

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

[0.10.2]: https://github.com/Forevka/canvas-word/releases/tag/v0.10.2
[0.10.0]: https://github.com/Forevka/canvas-word/releases/tag/v0.10.0
[0.8.0]: https://github.com/Forevka/canvas-word/releases/tag/v0.8.0
[0.7.5]: https://github.com/Forevka/canvas-word/releases/tag/v0.7.5
[0.7.4]: https://github.com/Forevka/canvas-word/releases/tag/v0.7.4
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
