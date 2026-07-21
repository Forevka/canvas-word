# Feature roadmap

Goal: a document editor where a Word user writes a real multi-page document and
doesn't notice it isn't Word. Every item below is shipped and checked off, with
links to the tests that verify it. Items marked _Playwright_ are UI-verified via
scripted browser checks (`shots/`) rather than a unit test.

Conventions:

- **op** = an entry in `shared/src/model/ops.ts` — returns an exact inverse +
  position mapper (the undo/collab invariant).
- **the walk** = the pagination loop in `frontend/src/layout/engine.ts`.
- **band machinery** = per-page header/footer layout with `{page}` substitution.
- Importer coordination assumes the docx prop tables in
  `frontend/src/import/docx/` (shared id spaces; the resolver flattens cascades
  into concrete runs, keeping `namedStyle`/list refs).

---

## Core editing & Word furniture

- [x] **Lists — bullets, numbered, multilevel (9 levels)**
  - Model: `Document.lists` (id space = docx `numId`); `ListDefinition.levels[0..8]` mirrors OOXML `numbering.xml` (`format`, `%N` text pattern, indent/hanging, `start`, `markerStyle`); `ParaStyle.list`; op `setListDefinition`.
  - Engine: numbering pass at the top of the walk (per-`(listId, level)` counters, deeper levels reset on increment, `start`/restart honored); markers are paint-only (`PlacedBlock.marker`), right-aligned in the hanging indent so the line cache stays valid.
  - Input: `toggleList`; Tab/Shift+Tab at offset 0 demote/promote; Enter inherits, Enter-on-empty exits; Backspace ladder (promote → leave list → merge).
  - Tests: [`lists.test.ts`](shared/src/model/lists.test.ts), [`listsLinks.test.ts`](frontend/src/import/docx/listsLinks.test.ts), [`bulletGlyph.test.ts`](frontend/src/paint/bulletGlyph.test.ts), [`listMarker.test.ts`](frontend/src/export/pdf/listMarker.test.ts)

- [x] **Find & replace**
  - Search walks `paragraphsOf(doc)` (body + cells + bands) collecting `{ blockId, start, end }`; options match-case / whole-word; no cross-paragraph matches.
  - UI: Ctrl+F bar (count badge, prev/next, replace + replace-all, Esc). Highlights ride a second paint rect channel (`setSearchRects`). Replace-all is ONE transaction, applied back-to-front so offsets stay valid.
  - Tests: [`documentEditor.test.ts`](shared/src/model/documentEditor.test.ts) (replace engine); Playwright (bar UI)

- [x] **Hyperlinks**
  - Model: `CharStyle.link` (joins `styleEq`, so runs split/merge at link boundaries). Paint: `#0b57d0` + underline unless overridden. Input: hover pointer + `title`, Ctrl+click opens (`window.open`, `rel=noopener`). Commands: `setLink` / `removeLink`. Clipboard maps `<a href>` both ways; importer maps `w:hyperlink` + rels.
  - Tests: [`listsLinks.test.ts`](frontend/src/import/docx/listsLinks.test.ts), [`clipboard.test.ts`](frontend/src/input/clipboard.test.ts)

- [x] **Highlight + subscript/superscript**
  - Model: `CharStyle.highlightColor`, `CharStyle.verticalAlign` (`sub`/`super`). `charStyleToFont` returns the scaled size (0.65×) so pretext measures what paint draws; paint adds the baseline shift. Highlight rects paint per fragment between the selection and text passes.
  - Tests: [`runEffectsPaint.test.ts`](frontend/src/paint/runEffectsPaint.test.ts), [`runEffectsRoundtrip.test.ts`](frontend/src/export/docx/runEffectsRoundtrip.test.ts)

- [x] **Format painter**
  - `editor.copyFormat()` captures `{ char, para }` into a one-shot slot; the next selection applies both in one transaction. Double-click = sticky until Esc.
  - Tests: Playwright

- [x] **AutoCorrect (typographic)**
  - Hook before `insertText`, triggered on word delimiters: straight→curly quotes (by left context), `--`→—, `(c)`→©, `(tm)`→™, optional capitalize-after-period (default off). Extra ops ride the same typing transaction, so they coalesce into one undo step.
  - Tests: Playwright

- [x] **Soft line break (Shift+Enter)**
  - Model: a `"\v"` sentinel inside run text = hard line break within the paragraph (1 UTF-16 unit, no caret/backspace special-casing). Engine: `paragraphLines` splits runs into `\v`-delimited segments, lays out each as its own pretext prepare+break; a `\v`-terminated line counts as paragraph-last (ragged under justify). Export/import: `w:br`; HTML clipboard `<br>` both ways.
  - Tests: [`engine.test.ts`](frontend/src/layout/engine.test.ts), [`clipboard.test.ts`](frontend/src/input/clipboard.test.ts)

## Export

- [x] **.docx export**
  - `frontend/src/export/docx/writeDocx.ts` emits `document.xml`, `styles.xml`, `numbering.xml`, `footnotes.xml`, `settings.xml`, header/footer parts, `[Content_Types].xml`, `.rels`, zipped with `fflate`. Mappings are exact inverses of the importer decoders (shared prop-table module prevents drift): runs → `w:r`/`w:rPr`, paragraphs → `w:p`/`w:pPr`, tables → `w:tbl` + `tblGrid` (`gridSpan`, re-synthesized `w:vMerge`), images → `w:drawing` (inline/anchored, EMU conversions), stylesheet → basedOn chains.
  - Runs carry full direct formatting so `w:pStyle` can't leak run props on re-import. Round-trip oracle `writeDocx → runImport → compare` holds block count (8.9 MB report: 1031 → 1031). Runs in a browser worker and on the Node backend. See [`EXPORT.md`](EXPORT.md).
  - Tests: [`docx.test.ts`](frontend/src/export/docx/docx.test.ts), [`roundtripStability.test.ts`](frontend/src/export/roundtripStability.test.ts), [`e2e.test.ts`](frontend/src/export/e2e.test.ts)

- [x] **PDF export (vector)**
  - `frontend/src/export/pdf/renderPdf.ts` re-runs the layout engine to get the `LayoutTree`, then replays each page into pdfkit — a constant-for-constant inverse of `paint/renderer.ts`, so a PDF page matches the canvas pixel-for-pixel. Bundled metric-clone faces subset-embed via fontkit over the shared DOM-free measure host, so editor = browser export = Node export paginate identically (see [`FONTS.md`](FONTS.md)). 8.9 MB / 81-page report ≈ 1.2 s. Raster browser-print was dropped in favor of this path.
  - Tests: [`pdf.test.ts`](frontend/src/export/pdf/pdf.test.ts), [`serverRenderPdf.test.ts`](backend/src/export/serverRenderPdf.test.ts)

## Sections, columns, fields

- [x] **Sections (page setup per part)**
  - Next-page breaks via `ParaStyle.sectionBreak` carrying a `SectionPatch` (absent fields inherit `doc.section`, so bands link-to-previous); mirrors OOXML sectPr-on-paragraph 1:1. The walk swaps page geometry (size, margins, bands) at each break. UI: `setSectionProps` op, §⏎ break, 📐 page-setup panel (Letter/A4/Legal × orientation × margins). Verified with mixed Letter + A4-landscape, per-section bands, typing on resized pages.
  - Tests: [`sectionLayout.test.ts`](frontend/src/editor/sectionLayout.test.ts), [`engine.test.ts`](frontend/src/layout/engine.test.ts), [`multiSection.test.ts`](frontend/src/import/docx/multiSection.test.ts)

- [x] **Newspaper columns**
  - `SectionProps.columns { count, gapPx }` (patch `null` = off vs absent = inherit). Flow: `newPage()` → next column → hard page; floats clamp to their column; tables/images measure at column width. Column break = Ctrl+Shift+Enter (`ParaStyle.columnBreakBefore`). Last-page balancing deferred (Word only balances continuous-section ends).
  - Tests: [`engine.test.ts`](frontend/src/layout/engine.test.ts), [`sectionLayout.test.ts`](frontend/src/editor/sectionLayout.test.ts)

- [x] **First / odd / even headers & footers**
  - Six band containers (`header`/`footer` + `First`/`Even` variants) as first-class Containers. Band layout picks first-of-section → even-number → default per page and records `Page.headerSource/footerSource`. `{page:roman|Roman|alpha|Alpha}` formats + per-section `pageNumberStart`; "Different first page" / "Different odd & even" seed/remove variant bands (`setSectionBand` op). Tall bands push the content box (`Page.contentTopPx/contentBottomPx`); band tables are editable.
  - Tests: [`imagesHeaders.test.ts`](frontend/src/import/docx/imagesHeaders.test.ts), [`sectionLayout.test.ts`](frontend/src/editor/sectionLayout.test.ts)

- [x] **Fields & Table of Contents**
  - Inline fields `Run.field { type: 'page'|'pages'|'date'|'time'|'ref', fmt? }` resolved by the generalized band substitution machinery. TOC entries are real paragraphs (`ParaStyle.tocEntry { targetId, level }`) from Heading 1–3 (☰§ inserts/regenerates). Page numbers are paint-only (`PlacedBlock.toc`), resolved post-pass against the final page map — always fresh, no fixpoint, since entries reserve a 48px right gutter. Dot leaders via `setLineDash`; Ctrl+click jumps to the heading. Deferred: body-text PAGE fields, pretext tab stops.
  - Tests: [`toc.test.ts`](shared/src/toc.test.ts), [`tocBuild.test.ts`](frontend/src/import/docx/tocBuild.test.ts), [`tocField.test.ts`](frontend/src/export/docx/tocField.test.ts), [`fieldCommands.test.ts`](frontend/src/editor/fieldCommands.test.ts), [`fields.test.ts`](shared/src/fields.test.ts)

- [x] **Footnotes**
  - `CharStyle.footnoteRef` (the ref run's text is its number, kept in sync by the insert command's renumber pass) + `Document.footnotes` + `setFootnote` op. Engine: a greedy single pass — a chunk carrying refs shrinks (`fitChunkWithNotes`) until chunk + notes co-fit above the dynamic `bottomY()`; reservations are per page (columns share the area). Notes lay out post-walk as real page blocks under a ⅓-width separator, editable in place. Deferred: refs in cells/floats, note splitting across pages, endnotes.
  - Tests: [`sectionFootnotes.test.ts`](frontend/src/import/docx/sectionFootnotes.test.ts), [`endnoteCommands.test.ts`](frontend/src/editor/endnoteCommands.test.ts)

- [x] **Break controls**
  - `ParaStyle.keepLinesTogether` (`w:keepLines`) — a paragraph that would split moves whole unless taller than a page. keep-with-next chains — every chained member must fit whole and the terminator must keep an orphan/widow-legal first take, else the group breaks before its first member. Importer maps `w:keepLines`/`w:keepNext`; `setParaProps` sets them via API.
  - Tests: [`engine.test.ts`](frontend/src/layout/engine.test.ts), [`pages.test.ts`](frontend/src/layout/pages.test.ts)

## Review layer

- [x] **Track changes & comments**
  - A review **overlay** (`shared/src/review/` + `frontend/src/review/`) the OOXML core never sees, with a three-way mode switch (Editing / Suggesting / Viewing). In suggest mode a transaction interceptor rewrites destructive edits into non-destructive attributed records (inserts → author-colored underlines, deletes → strikethroughs, format → margin change-bars); anchors ride the same `mapPosition` rebasing as bookmarks, with a GC pass for emptied records.
  - Threaded comments anchor to a range with a floating composer + docked Review pane (Accept/Reject, accept/reject-all, click-to-scroll). @-mentions from an embedder `knownUsers` roster fire an HMAC-signed `comment.mention` webhook. Accept/reject emit ordinary core ops; everything syncs on an idempotent, causally-delivered review channel (Postgres `review_ops`) and rehydrates on join. Default `.docx`/PDF export bakes the overlay (reject-all = original baseline). Tracks insert/delete/format plus structural edits (split/merge, block & table row/col ops) and multi-op paste / cross-paragraph deletes (decomposed per-op); rich comment bodies remains. See [`REVIEW.md`](REVIEW.md).
  - Tests: [`intercept.test.ts`](frontend/src/review/intercept.test.ts), [`resolve.test.ts`](frontend/src/review/resolve.test.ts), [`ops.test.ts`](shared/src/review/ops.test.ts), [`rebase.test.ts`](shared/src/review/rebase.test.ts), [`rehydrate.test.ts`](shared/src/review/rehydrate.test.ts), [`SyncClient.review.test.ts`](frontend/src/sync/SyncClient.review.test.ts)
