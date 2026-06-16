# Headless TOC & field rendering (PDF)

How `runExport(doc, "pdf")` — and the stateless backend route `POST /render.pdf`
(`backend/src/export/serverExport.ts → renderPdfFromDocx`) — resolve the
layout-derived things a renderer-less producer (e.g. a C# pipeline emitting raw
`.docx`) can't compute: the table of contents and page-number fields.

The page number of a TOC entry and of a `PAGE`/`NUMPAGES` field are **paint-only**:
they are resolved by the layout engine on every relayout and painted, never stored
in the model. So they are *always* current in the rendered PDF.

## PAGE / NUMPAGES (footer "Page X of Y")

Header/footer `PAGE`/`NUMPAGES` fields import as live `{page}`/`{pages}` tokens
(`import/docx/documentParser.ts`). The layout substitutes them **per page** on
per-page footer/header clones (`layout/engine.ts`, `substituteTokens`), so every
page shows the correct number. Nothing to configure; nothing can go stale.

## Table of contents

A `TOC` field is handled by `generateTocIntoDoc(doc, tocOptions)`
(`@cw/shared` → `shared/src/toc.ts`), called by `renderPdfFromDocx` when the
document has a TOC field. There are two cases:

### 1. Empty TOC field → BUILD from headings
When the field has no entries (e.g. C# emits a `TOC` field whose result is a
"Please press F9…" placeholder), import records the field's block as
`Document.tocAnchorBlockId` (captured at parse/map time — see "Anchor" below).
`generateTocIntoDoc` then builds entries from the document's headings and splices
them at that block. The built entries:

- **Honor the field's switches** — `\o "1-5"` sets the level range, `\t` maps custom
  styles (`parseTocInstruction`).
- **Inherit the document's own TOC styling** — the entry base style is taken from the
  TOC field paragraph's resolved style (its `pStyle`, e.g. `TOC 1`), so the result
  matches the source Word doc. `TocOptions` passed to `/render.pdf` override on top.
- **Omit the auto-title** when filling a field (the doc usually has its own heading).

Result: passing `{}` (or no `toc` part) reproduces the source's intended TOC.

### 2. Populated TOC field → PRESERVE
When the field already has entries (a Word/C#-rendered TOC), import marks them as
`tocEntry` and strips their cached page numbers (`markImportedTocEntries`).
`generateTocIntoDoc` **leaves them untouched** — it does NOT rebuild from headings.
The entries keep the source's own per-level styling, and layout repaints their page
numbers. (Rebuilding would clobber the author's TOC with our generator defaults.)

## Staleness semantics (populated TOC) — by design

For a **populated** TOC we preserve the imported entry **set** and **label text**
verbatim; only **page numbers** self-heal. We do **not** reconcile entries against the
document's current headings. So the PDF reflects exactly what the editor currently
shows (imported entries + live numbers), **not** a forced "Update TOC / F9":

| Change in the document | Effect in the rendered PDF |
|---|---|
| A heading moved to another page | ✅ Page number is recomputed (correct). |
| A heading was **deleted**, entry lingers in the TOC | Entry **stays**, but its target no longer resolves → it renders with **no page number** (label only). Not removed. |
| A heading was **added** after the TOC was built | ❌ **Not listed** — we don't inject new entries. |
| A heading was **renamed** | ❌ Entry keeps its **old label**; page number still correct. |

Only the **empty-field BUILD path** is fully fresh (entries derived from current
headings). A future "refresh a populated TOC" option would need to sample the
existing entries' per-level styling (cf. the editor's `tocOptionsFromExisting`)
before rebuilding, to avoid regressing the look — which is why *preserve* is the
safe default.

## What is NOT eligible to be a TOC entry

Heading detection (`detectTocHeadings` → `bodyParagraphs`) considers **only
paragraph blocks** (recursing into table cells). Therefore:

- **Tables** (`TableBlock`) and **images** (`ImageBlock`) can never become TOC
  entries — they aren't paragraphs and the model has no paragraph style on them.
- An **image-only paragraph styled as a heading** yields empty label text
  (`textOfRuns` ignores drawings) and is **skipped** (`buildTocParagraphs` drops
  empty-text entries).
- A heading-styled **paragraph inside a table cell** *is* listed — as a normal text
  entry (its text), which is correct; the table itself is not embedded.

So unlike some engines (e.g. Syncfusion, which can embed a heading-styled image or
table directly into the TOC entry), a misapplied heading style on an image/table
here never pulls that object into the contents — entries are always plain text.

## Anchor (why placement is reliable)

The empty-field anchor (`tocAnchorBlockId`) is captured **during parsing/mapping**,
not by counting `<w:p>` elements: `documentParser` tags `IRParagraph.tocField` when a
paragraph holds a `TOC` instruction, and `mapToModel` records the actual model block
that paragraph produced (`mapper.tocField()`). This is robust because `mapParagraph`
can split one `<w:p>` into several model blocks, and cover-page `<w:sdt>`s / tables /
drawings make any raw-XML ordinal diverge from the model index.
