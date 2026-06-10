# Feature roadmap

The goal: a document editor where a typical Word user writes a real multi-page
document and doesn't notice it isn't Word. This was the working plan to get there
— the gaps a user hits in the first five minutes, then the document-level gaps
that follow. **Everything below has shipped.** Each entry records the
model/engine/input/UI deltas, undo story, edge cases, and verification.
Conventions referenced throughout:

- **"op"** = an entry in `model/ops.ts` — must return an exact inverse + position
  mapper (the undo/collab invariant).
- **"the walk"** = the pagination loop in `layout/engine.ts#layoutDocument`.
- **"band machinery"** = per-page header/footer layout with `{page}` substitution.
- Importer coordination notes assume the docx workstream's prop tables in
  `frontend/src/import/docx/` (shared id spaces; their resolver flattens cascades
  into concrete runs and keeps `namedStyle`/list refs).

---

### Lists (bullets, numbered, multilevel) — ✅ DONE
**Shipped as planned; verified: counter sequence 1./2./a./b./i./3. with level
resets, hanging-indent markers, toggle, Tab ladder, Enter semantics, undo.**

Model
- `Document.lists?: Map<string, ListDefinition>` (id space = docx `numId`).
- `ListDefinition.levels[0..8]`: `{ format: 'bullet'|'decimal'|'lowerLetter'|'upperLetter'|'lowerRoman'|'upperRoman'; text: string /* "%1." pattern */; bulletChar?: string; indentLeftPx: number; hangingPx: number; start: number; markerStyle?: Partial<CharStyle> }` — a direct mirror of OOXML `numbering.xml` abstractNum levels, so the importer feeds it losslessly.
- `ParaStyle.list?: { listId: string; level: number }`.
- New op `setListDefinition { listId, def }` (inverse = old def), plus list
  membership changes ride the existing `setParaStyle` op.

Engine
- **Numbering pass** at the top of the walk (document order, body only):
  maintain counters per `(listId, level)`; incrementing level N resets all
  deeper levels; honor `start` and restart-after-gap rules. Output: marker
  string per list paragraph. Pure string arithmetic — no caching needed; runs in
  the measured[] pass.
- Markers are **paint-only**: `PlacedBlock.marker?: { text, style, x }` — drawn
  right-aligned in the hanging indent (`indentLeftPx − hangingPx … indentLeftPx`),
  never measured by pretext (text x-origin is `indentLeftPx`, so the line cache
  stays valid; the marker can't change line breaking, same as Word's overlap
  behavior for very long markers).
- Counters depend on *preceding* paragraphs → marker recompute is part of every
  walk (cheap), but LineBox caches are untouched.

Input / commands
- `toggleList('bullet' | 'decimal')` on the selection's paragraphs (assign to a
  fresh or nearest-matching listId).
- **Tab / Shift+Tab with caret at offset 0 of a list paragraph** = demote /
  promote level (checked BEFORE table-cell Tab navigation in the key router).
- **Enter** inherits the list membership (split keeps `ParaStyle.list` — already
  free, split clones style); **Enter on an empty list paragraph exits the list**
  (Word). **Backspace at start**: level > 0 → promote; level 0 → leave list;
  then the normal pageBreak/merge ladder.
- Toolbar: bullet + numbered buttons; style gallery styles may carry list refs
  (ListParagraph style).

Verify
- Counter sequences across levels incl. resets; `start` offsets; two interleaved
  lists keep independent counters; Enter/Tab/Backspace ladder; undo each;
  markers render in hanging indents at all alignments; stress doc unaffected.

Importer coordination: `numbering.xml → Document.lists`, `w:numPr → ParaStyle.list`.

### Find & replace — ✅ DONE
**Shipped as one batch with hyperlinks, highlight, sub/superscript, format painter and autocorrect.**

- Search model: walk `paragraphsOf(doc)` text (body + cells + bands); collect
  `{ blockId, start, end }[]`. Plain text, options: match case, whole word.
  No cross-paragraph matches (Word's plain search doesn't either).
- UI: Ctrl+F floating bar (app chrome, like the toolbar): input, count badge
  "3/17", prev/next (Enter / Shift+Enter), replace + replace-all fields, Esc closes.
- Highlights: paint layer gets a second rect channel `setSearchRects(rects)`
  (orange, under selection) — same per-page invalidation as selection rects;
  current match uses the existing selection.
- Replace: `deleteRange` + `insertText` at the match (inherits left style);
  **replace-all is ONE transaction** → single undo. Offsets of later matches
  shift — apply back-to-front to keep them valid.
- Verify: matches in cells and headers (band scope jump on navigate enters story
  mode), replace-all single-undo, case/whole-word, zero matches.

### Hyperlinks — ✅ DONE

- Model: `CharStyle.link?: string` (must join `styleEq` so runs split/merge
  correctly at link boundaries).
- Paint: linked fragments default to `#0b57d0` + underline unless overridden.
- Input: hover over a linked fragment → pointer cursor + native `title` tooltip
  on the container; **Ctrl+click opens** (`window.open`, rel=noopener); plain
  click just places the caret (Word).
- Commands: `setLink(url)` on a range (collapsed caret expands to the word),
  `removeLink()`; toolbar 🔗 button prompting for URL.
- Clipboard: `fragmentToHtml` wraps linked runs in `<a href>`; `htmlToFragment`
  maps `<a>` back. Importer: `w:hyperlink` + rels → link prop (their side).
- Verify: split/merge at link edges, copy→paste round-trip, ctrl+click, undo.

### Highlight + subscript/superscript — ✅ DONE

- Model: `CharStyle.highlightColor?: string`, `CharStyle.verticalAlign?: 'sub'|'super'`.
- Measurement stays honest: `charStyleToFont` returns the **scaled** size
  (0.65×) for sub/super so pretext measures what paint draws; paint adds the
  baseline shift (−0.38em super / +0.16em sub) at fillText time only.
- Paint: highlight rects per fragment between the selection pass and text pass.
- Toolbar: highlight swatch row, x² / x₂ toggles (three-state with normal).
- Verify: mixed sub/super/normal in one line keeps a stable baseline; caret
  height on scaled runs; highlight under selection color blend.

### Format painter — ✅ DONE

- `editor.copyFormat()` captures `{ char: styleAtRuns(caret), para: caretBlock.style }`
  into a one-shot slot; cursor swaps to a brush; next selection (mouseup with a
  range, or click = word) applies char via `setCharStyle` + para via
  `setParaProps` in one transaction. Double-click the button = sticky until Esc
  (Word). Toolbar 🖌 button.

### AutoCorrect (typographic) — ✅ DONE

- Hook in the wiring layer before `insertText` dispatch, triggered on
  word-delimiters (space, punctuation, Enter): straight→curly quotes
  (open/close by left context), `--`→—, `(c)`→©, `(tm)`→™, optional
  capitalize-after-period (default OFF).
- Implemented as extra ops inside the same typing transaction → coalesces into
  the typing undo step. (Word reverts the correction alone on first Ctrl+Z;
  acceptable divergence, noted.)
- Settings object on the editor so the demo can toggle rules.

### Soft line break (Shift+Enter) — ✅ DONE

- Model: a **vertical-tab sentinel** `"\v"` inside run text = hard line break
  within the paragraph (1 UTF-16 unit → offsets/caret/backspace need zero
  special-casing; it's just a character).
- Engine: `paragraphLines` splits the paragraph's run list into `\v`-delimited
  segments, lays out each segment as its own pretext prepare+break sequence,
  concatenates the LineBoxes (the `\v` belongs to the end of its segment's last
  line — offset affinity like eaten whitespace). Line cache key unchanged
  (revision covers it). Justify: a `\v`-terminated line counts as paragraph-last
  (stays ragged, Word behavior).
- Input: Shift+Enter in the key router → `insertText("\v")`; the proxy's
  `insertLineBreak` beforeinput maps to the same.
- Export/import: `w:br` (their side); HTML clipboard `<br>` both directions.
- Verify: break mid-styled-run, caret/Home/End treat segments as separate visual
  lines (they ARE separate LineBoxes — free), backspace deletes the sentinel.

---

### .docx export — ✅ DONE
**Shipped: `frontend/src/export/docx/writeDocx.ts` emits `document.xml`, `styles.xml`,
`numbering.xml`, `footnotes.xml`, `settings.xml`, header/footer parts,
`[Content_Types].xml` and `.rels`, zipped with `fflate`. Runs carry full direct
formatting (every toggle explicit) so `w:pStyle` can't leak run props on re-import;
row spans re-synthesize the `w:vMerge="continue"` cells the importer drops; list ids
remap to integer `numId`s. Correctness oracle: `writeDocx → runImport → compare` —
the 8.9 MB report re-imports to the exact same block count (1031 → 1031). Runs both in
a browser worker and on the Node backend. See `EXPORT.md`.**

- Emit with `fflate`: `[Content_Types].xml`, `_rels/.rels`,
  `word/document.xml`, `word/styles.xml`, `word/numbering.xml` (after lists),
  `word/header1.xml` / `footer1.xml` + sectPr refs, `word/media/*` +
  `word/_rels/document.xml.rels`.
- Mapping (exact inverses of the importer's decoders — share one prop table
  module to prevent drift):
  - runs → `w:r`/`w:rPr` (b, i, u, strike, sz half-points, rFonts, color,
    spacing, vertAlign, highlight, hyperlink wrapping)
  - paragraphs → `w:p`/`w:pPr` (jc, spacing, ind incl. firstLine/hanging,
    keepNext, pageBreakBefore, pStyle, numPr)
  - tables → `w:tbl` + `tblGrid` from `colFractions × printable width`,
    `gridSpan` for colSpan, nested content recursively
  - images → inline `w:drawing` (block) / anchored with wrap squares (float),
    EMU conversions, media parts; data-URI images decoded to bytes
  - stylesheet → styles.xml with basedOn chains
- Acceptance: **round-trip** `import(export(doc))` equals doc on a normalized
  comparison, and the file opens clean in actual Word/LibreOffice.

### Print & PDF export — ✅ DONE (vector PDF)
**Shipped the vector path directly (skipping the raster phase): `frontend/src/export/pdf/renderPdf.ts`
re-runs the editor's own layout engine to get the `LayoutTree`, then replays each page
into **pdfkit** (not pdf-lib as planned here) — a constant-for-constant inverse of
`paint/renderer.ts`, so a PDF page matches the canvas pixel-for-pixel. Fonts: the bundled
metric-clone faces are subset-embedded via fontkit (no system fonts), the same DOM-free
measure host the editor and the Node backend share, so editor = browser export = Node
export paginate identically (8.9 MB report = 80 pages everywhere; see `FONTS.md`). Browser
runtime verified end-to-end via Playwright; the Node path is fully test-covered.
Real 8.9 MB / 81-page report: PDF ~1.2 s. See `EXPORT.md`.**

- **Phase A (raster)**: print stylesheet that hides chrome and lays each page
  canvas at exact physical size (`@page { size: letter; margin: 0 }`); force
  re-render of ALL pages at print DPR (no virtualization during print). Quick,
  universally correct, raster text.
- **Phase B (vector)**: replay the LayoutTree into `pdf-lib` — same fragments,
  same coordinates, real text. The blocker is fonts: system fonts can't be
  extracted, so bundle subsettable webfonts (e.g. PT Serif/Sans + JetBrains
  Mono) as the default doc fonts, embed subsets via fontkit. Images decode from
  data-URIs/URLs into XObjects. Headers/footers/tables/floats come free — they
  are already absolute geometry in the tree.

### Sections (page setup per part) — ✅ DONE
**Shipped: next-page section breaks (`ParaStyle.sectionBreak` carrying a
`SectionPatch` — absent fields inherit doc.section, so bands "link to
previous"), per-page Page dims in the tree, per-section measure widths +
markers, per-page placeholders/canvases/clientToPage (cumulative offsets),
`setSectionProps` op for the final section, §⏎ toolbar break + 📐 page-setup
panel (Letter/A4/Legal × orientation × margin presets) targeting the caret's
section. Verified: Letter + A4-landscape mixed, per-section bands & margins,
caret/typing on resized pages, undo/redo, stress clean (1217pp warm 4.9ms).**

- Model: section-break paragraphs — `ParaStyle.sectionBreak?: { type: 'nextPage'; props: SectionProps }`
  (OOXML puts sectPr ON a paragraph; mirroring that keeps import/export 1:1 and
  avoids restructuring `Document.blocks`). The document-final `doc.section`
  stays as the last section's props.
- Engine: the walk tracks current section props; a section break forces
  `newPage()` and swaps page geometry (size, margins, bands). Page placeholders
  get per-page dimensions (paint already reads sizes from the tree — make
  width/height per-Page instead of per-tree).
- `continuous` sections deferred until columns exist (that's their only point).
- UI: "Insert section break (next page)" + a page-setup dialog (size presets,
  orientation, margins) applying to the caret's section.
- Verify: A4 + Letter mixed, landscape, per-section headers, undo.

### Newspaper columns — ✅ DONE
**Shipped: `SectionProps.columns {count,gapPx}` (SectionPatch uses `null` =
explicitly off vs absent = inherit), flow `newPage()` → next column → hard page;
all "empty page" guards became "empty column" (colHasContent), floats clamp to
their column, list-marker x deferred to placement time, tables/images measure
at column width. Ctrl+Shift+Enter → `ParaStyle.columnBreakBefore` with the
Word backspace ladder (break clears before merge); Columns select in 📐 panel.
Verified: 2-col layout with tables/lists/images, column break lands at next
column top, arrow nav across the boundary, undo, stress 0 violations.**

- `SectionProps.columns?: { count: number; gapPx: number }`.
- Engine: the page content box divides into N column boxes; `newPage()`
  generalizes to `nextColumn()` (advance column, page-break when columns are
  exhausted). Everything downstream already works on absolute coords —
  geometry, paint, selection untouched. Floats clamp to their column; tables
  wider than a column span... (Word: tables can overflow the column — match).
- Column break = Ctrl+Shift+Enter → `ParaStyle.columnBreakBefore`.
- Last-page column balancing: defer (Word only balances continuous-section ends).
- Verify: 2/3-col layouts, breaks, caret Up/Down across column boundaries
  (line-index order already follows fill order).

### First/odd/even headers & footers — ✅ DONE
**Shipped: six band containers (`header/footer` + `First/Even` variants) as
first-class Containers (ops, locator, paragraphsOf); band layout picks
first-of-section → even-displayed-number → default per page and records
`Page.headerSource/footerSource`; story-edit navigates the variant governing
THAT page; `{page:roman|Roman|alpha|Alpha}` formats + per-section
`pageNumberStart` (panel "Number from"); "Different first page"/"Different odd
& even" checkboxes seed/remove empty variant bands (`setSectionBand` op).
Verified: variant pick incl. restart-at-5 parity, editing headerFirst leaves
the default header untouched, undo, 77 tests, stress warm 6.3ms.
Along the way (user-reported): tall bands now PUSH the content box (Word) —
`Page.contentTopPx/contentBottomPx`; band-edit dash/dim follow it; band
TABLES are fully editable (imported footers: table + page-number paragraph).**

### Fields & Table of Contents — ✅ DONE
**Shipped: band field tokens extended with `{date}`/`{time}` (+ the header
`{page:fmt}` formats); TOC as REAL paragraphs tagged `ParaStyle.tocEntry`
{targetId, level} generated from Heading1–3 paragraphs (☰§ button inserts at
the caret or regenerates in place). Page numbers are PAINT-ONLY decorations
(`PlacedBlock.toc`) resolved in an engine post-pass against the final page map
— always fresh, zero fixpoint, because entries are measured with a reserved
48px right gutter so the number can't affect line breaking. Dot leaders drawn
with setLineDash (no tab stops needed); wrapped entries put the number on
their last line. Ctrl+click an entry jumps the caret to its heading.
Deferred: body-text PAGE fields (bands only — documented), bookmarks/`ref`
fields, pretext tab stops.**

- Inline fields: `Run.field?: { type: 'page'|'pages'|'date'|'time'|'ref'; fmt?: string }`
  — rendered by the band substitution machinery generalized to a field resolver;
  body-text PAGE fields restricted to bands initially (the per-page re-prepare
  cost is why bands re-lay per page; body fields would need the same treatment —
  documented limitation until someone needs it).
- **TOC**: block-level generated region: scan heading-styled paragraphs after
  layout, build entries (text + page number from the tree), regenerate the TOC
  blocks, relayout once (TOC length changes page numbers → iterate to fixpoint,
  converges in ≤2 passes in practice; cap at 3). Entries are real paragraphs
  with right-aligned tab leaders (needs tab stops: pretext exposes
  `tabStopAdvance` — wire `ParaStyle.tabStops` through prepare options).
- Click a TOC entry → scroll to target (bookmarks: `Run.bookmark?: string`).

### Footnotes — ✅ DONE
**Shipped: `CharStyle.footnoteRef` (the ref run's TEXT is its number, kept in
sync by the insert command's renumber pass — model text and area markers can
never disagree) + `Document.footnotes: Record<id, Paragraph[]>` + `setFootnote`
op. Engine: a GREEDY single pass instead of the per-page fixpoint — a chunk
carrying refs shrinks (`fitChunkWithNotes`) until chunk + notes co-fit above
the dynamic `bottomY()` (= contentBottom − reserved); reservations are per
PAGE (columns share the area); the area lays out post-walk as REAL page blocks
under a ⅓-width separator rule, so notes are caret/select/type-editable IN
PLACE with zero story-scope machinery (notes joined ParaLocation /
split / merge instead). a¹ toolbar button; caret drops into the new note
(Word). Verified: insert-before renumbers 1↔2, superscript refs, in-place note
editing, growing a note pushes body lines off the page with the rule invariant
holding on every page, undo, stress clean.
Deferred: refs inside table cells/floats (top-level body paragraphs only),
note splitting across pages (overflow), endnotes.**

### Remaining break controls — ✅ DONE
**Shipped: `ParaStyle.keepLinesTogether` (w:keepLines) — a paragraph that
would split moves whole unless taller than a page; keep-with-next CHAINS —
every chained member must fit whole and the chain terminator must keep its
orphan/widow-legal first take, else the whole group breaks before the first
member (closes the pairwise TODO). Verified: 6-line paragraph 3+2 split → moves
whole; h1→h2→body group moves together where pairwise stranded the headings;
stress: 495 keepNext headings, 0 stranded, warm 5.8ms. (No toolbar UI —
importer maps w:keepLines/w:keepNext; `setParaProps` sets them via API.)**

---

## Status

Every feature above has shipped — including `.docx` and PDF export (see
`EXPORT.md`). The only item dropped from the original plan is raster
browser-print, covered by the vector PDF path. The editor covers ~95% of everyday
Word usage. Work since then has moved past this plan into collaboration, a Node
backend, an admin dashboard, and npm distribution — see `README.md`.
