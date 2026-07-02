# Merge Documents — implementation plan

Status: **PLAN** (not yet started). Goal: a first-class **merge / append** operation that
concatenates two (or N) WordCanvas documents into one — fully in the model — so headless
assembly pipelines can fold rendered parts together without a round-trip through a
word-processor library. It lands on all three surfaces: pure `@cw/shared` core →
`@forevka/wordcanvas` (npm) → C# ClearScript bindings (`WordDocument.Append` / `engine.Merge`).

Related: [query-edit-api], `builder-mutator-additions.md` (backlog).

---

## 1. Why

A common headless pipeline renders sections separately, then assembles a destination
document: import each part, bring its content in under the destination's styles, control the
section break between parts, attach a TOC, and update fields. WordCanvas already has: import,
export, builder, TOC (layout-resolved page numbers — no reopen), header/footer bands, and a
37-op `applyOp` engine. **The only missing primitive is "append document B's content into
document A"** with correct id-space reconciliation. That is this plan.

Wins: order-of-magnitude less allocation, layout-resolved TOC/PAGE with no field-update
round-trip, and one pure function that is trivially unit-testable.

---

## 2. The model reality that makes merge non-trivial

The model is **concrete** (formatting baked on runs) but has many **keyed registries** and
**cross-references by id**. Naive block concatenation would collide ids and alias content. A
correct merge appends `source.blocks` after `dest.blocks` while **rebasing every id space**:

| Space | Where the ids live | Reference sites to rewrite |
|---|---|---|
| Block / cell ids | `Block.id`, `TableCell.id` | undo keys, layout cache, `bookmarks[*].start/end.blockId`, `tocEntry.targetId` |
| Named styles | `Document.stylesheet` | `ParaStyle.namedStyle`, `CharStyle.charStyleId` |
| List definitions | `Document.lists` (numId space — **will** collide) | `ParaStyle.list.listId` |
| Table styles | `Document.tableStyles` | `TableBlock.styleId` |
| SDT (content controls) | `Document.sdts` | `Block.sdtPath[]`, `CharStyle.sdtPath[]` (both, nesting-aware) |
| Fields | `Document.fields` | `Block.fieldId`, `CharStyle.fieldId` |
| Footnotes / endnotes | `Document.footnotes` / `.endnotes` | `CharStyle.footnoteRef` / `.endnoteRef` |
| Bookmarks | `Document.bookmarks` (keyed by **name**) | in-doc `#name` links, cross-references |
| Media | `ImageBlock.mediaId` (**content-addressed sha256**) | `ImageBlock.src`; the C# `cw-media:N` map |

Notes:
- **Media dedupes for free** — `mediaId` is `sha256(bytes)`, so identical logos/images across
  parts collapse to one id. Only the C#-side `cw-media:N → Uint8Array` map needs a union.
- **Bookmarks key by name** — collisions get renamed (`_Part_1` → `_Part_1__2`) unless the caller
  opts out; positions still need their block ids rebased. This handles duplicate bookmark
  ids/names across parts.
- **Style reconciliation has two modes** (use-destination vs keep-source), see §3.

---

## 3. Core API — `@cw/shared/model/merge.ts` (new)

```ts
export interface MergeOptions {
  /** Reconcile named styles / lists / table styles that collide with the destination.
   *  "useDestination" (default): a source style whose NAME matches a destination style is
   *  dropped and its references repointed at the destination's; source-only styles are added
   *  (renaming the id if the id — not the name — collides). Concrete run formatting is
   *  untouched either way (the model is concrete), so text keeps its look; this only decides
   *  which style ENTITY the round-trip reference points at.
   *  "keepSource": always keep the source's styles, renaming on id collision, so the source's
   *  own style definitions win for its content. */
  styles?: "useDestination" | "keepSource";              // default "useDestination"

  /** The section seam inserted between the destination's last section and the source's first.
   *  "nextPage"/"evenPage"/"oddPage" start the source on a new (parity-forced) page;
   *  "continuous" flows it inline in the same section geometry; "none" appends with no break
   *  (the caller manages breaks). Default "nextPage". */
  sectionBreak?: "nextPage" | "evenPage" | "oddPage" | "continuous" | "none";  // default "nextPage"

  /** Rename a source bookmark whose name already exists in the destination instead of dropping
   *  it. Default true. */
  renameBookmarksOnCollision?: boolean;                   // default true
}

export interface MergeResult {
  doc: Document;                       // NEW document (structural sharing; inputs untouched)
  idMap: MergeIdMap;                   // old→new per space (blocks, styles, lists, tableStyles,
                                       //   sdts, fields, footnotes, endnotes, bookmarks)
  warnings: MergeWarning[];            // e.g. renamed bookmark, dropped duplicate style
}

/** Append `source` after `dest`, reconciling every id space. Pure; returns a new doc. */
export function mergeDocuments(dest: Document, source: Document, opts?: MergeOptions): MergeResult;

/** Fold N docs left-to-right: mergeAll([a,b,c]) === merge(merge(a,b),c). */
export function mergeAll(docs: Document[], opts?: MergeOptions): MergeResult;
```

### The section seam (the subtle part)

`Document.section` is the FINAL/body `sectPr`; mid-document sections live as
`ParaStyle.sectionBreak` on the paragraph that ENDS them. To append source after dest:

1. Turn the destination's body section into a real section break: set
   `sectionBreak = { type, props }` (carrying dest's current `Document.section` geometry +
   header/footer bands) on the destination's **last block** (append a trailing empty paragraph
   first if the last block can't carry it — an image/table/equation). `type`/props come from
   `opts.sectionBreak`; `"continuous"` reuses dest geometry with no page break; `"none"` skips
   this step.
2. Append the (fully id-remapped) source blocks.
3. The source's own `Document.section` becomes the merged doc's new `Document.section` (its bands
   and geometry govern everything after the seam) — unless `"continuous"`, which keeps dest's.

This covers the common new-page vs continuous section strategies and removes the need for
callers to hand-manage a trailing section-boundary paragraph.

### Document-level scalars

`defaultTabStopPx`, `compatSettings`, `tocInstruction`/`tocAnchorBlockId`: **destination wins**;
source values are dropped (with a warning if they differ) — the merged doc is "the destination,
extended". A caller who wants a fresh TOC uses the builder's `tableOfContents()` after merge.

---

## 4. Edit-facade + op surface

- **`DocumentEditor.append(source: Document, opts?: MergeOptions): void`** — runs `mergeDocuments`
  against `this.doc` and replaces it as ONE undo step (snapshot inverse; too broad for a granular
  op, so it uses a `setDocument`-style replace op or a coarse snapshot inverse — decide in Phase 3).
- **`DocumentEditor.setSectionBand(sectionIndex, band, blocks): void`** — set a header/footer band
  on a specific section (by section index). Built on the EXISTING ops: `setSectionBand` for the
  final/body section, and `setParaStyle` patching `sectionBreak.props.<band>` for a mid-document
  section's break paragraph. This is what powers a **post-merge** footer pass (map
  bookmarks/markers → section index → assign a footer). Convenience wrappers
  `setSectionFooter` / `setSectionHeader`.

No NEW low-level op is strictly required for §3 (append rebuilds the block list + registries and
can ride a snapshot/replace op); the section-band setter reuses `setSectionBand` + `setParaStyle`.

---

## 5. npm + C# surface

- **npm** (`@forevka/wordcanvas`): export `mergeDocuments` / `mergeAll` (new `./merge` subpath, or
  fold into the existing `./query` edit exports — decide in Phase 3), plus `DocumentEditor.append`.
  Hand-written `types/merge.d.ts` + a `merge.parity.ts` compile guard, like the query surface.
- **C#** (`WordDocument` / `WordCanvasEngine`):
  - `WordDocument WordDocument.Append(WordDocument other, MergeOptions? opts = null)` — returns a
    new handle; **unions the two `cw-media` maps** so the appended images resolve on export.
  - `WordDocument WordCanvasEngine.Merge(params WordDocument[] docs)` / `Merge(IEnumerable<…>, MergeOptions)`.
  - `WordDocumentEditor.SetSectionFooter(int sectionIndex, Action<StoryBuilder> build)` /
    `SetSectionHeader(…)` for the post-merge pass.
  - A `MergeOptions` record (styles mode / section break / rename-bookmarks) with `ToJs`.
  - Backed by a `mergeBridge` JS fn on the ClearScript entry; covered by the existing
    `csharpBridgeParity.test.ts` guard. Requires a **bundle rebuild**
    (`node frontend/scripts/build-clearscript.mjs`).

---

## 6. Phased plan (one PR per phase; CodeRabbit between)

1. **Core append + id remap + style modes** (`@cw/shared/model/merge.ts`) — blocks/cells, styles,
   lists, table styles, sdts, fields, notes, bookmarks, media dedupe. Round-trip tests: build A & B
   with overlapping ids/styles/lists/bookmarks/sdts/footnotes → merge → assert counts, no id
   collisions, every reference resolves, both style modes behave.
2. **Section seam + break kinds + bands** — dest body → break paragraph, source section adoption,
   `continuous`/`none`. Tests: page-count and band-ownership after each break kind; export→reimport
   parity.
3. **npm surface** — `mergeDocuments`/`mergeAll` + `DocumentEditor.append` + types + parity guard.
4. **C# bindings** — `Append`/`Merge` + media union + `MergeOptions` + bundle rebuild + showcase demo
   (merge two built docs, export). Parity guard green.
5. **`setSectionFooter`/`setSectionHeader`** editor methods (TS + C#) for the post-merge footer pass.
6. **Examples + docs** — a `examples/merge-docs` demo, the footer recipe (§8), README section,
   CHANGELOG, memory update.

Effort: comparable to the SDT surface. Highest-risk = §1 remap completeness and §2 seam; both are
pure and heavily testable.

---

## 7. Usage example — merging documents

### TypeScript / npm

```ts
import { DocumentBuilder } from "@forevka/wordcanvas/builder";
import { mergeAll, mergeDocuments } from "@forevka/wordcanvas/merge"; // (subpath TBD — Phase 3)
import { exportDocx } from "@forevka/wordcanvas/export";

// Three independently-authored (or independently-imported) parts:
const cover   = new DocumentBuilder().paragraph("Cover", (p) => p.withStyle("Heading1")).build();
const summary = new DocumentBuilder().paragraph("Summary", (p) => p.withStyle("Heading1")).build();
const details = new DocumentBuilder().paragraph("Details", (p) => p.withStyle("Heading1")).build();

// Fold them left-to-right; each seam starts a new page and uses the destination's styles.
const { doc, warnings } = mergeAll([cover, summary, details], {
  sectionBreak: "nextPage",
  styles: "useDestination",
});
if (warnings.length) console.warn(warnings);

// Or step-by-step, inspecting the id remap between merges:
const step = mergeDocuments(cover, summary, { sectionBreak: "continuous" });
console.log(step.idMap.bookmarks); // { "_Part_1": "_Part_1__2", ... }

const bytes = await exportDocx(doc);   // TOC/PAGE resolve at layout — no reopen needed
```

Or via the stateful editor (undoable):

```ts
import { DocumentEditor } from "@forevka/wordcanvas/query"; // edit facade

const editor = new DocumentEditor(cover);
editor.append(summary, { sectionBreak: "nextPage" });
editor.append(details, { sectionBreak: "nextPage" });
editor.undo();                          // removes the details part in one step
const merged = editor.doc;
```

### C#

```csharp
using var engine = new WordCanvasEngine();

WordDocument cover   = engine.ImportDocx(File.ReadAllBytes("cover.docx"));
WordDocument summary = engine.ImportDocx(File.ReadAllBytes("summary.docx"));
WordDocument details = engine.ImportDocx(File.ReadAllBytes("details.docx"));

// One shot:
WordDocument report = engine.Merge(cover, summary, details);

// Or explicit, per-seam control:
report = cover
    .Append(summary, new MergeOptions { SectionBreak = SectionBreakKind.NewPage,
                                        Styles = StyleMergeMode.UseDestination })
    .Append(details, new MergeOptions { SectionBreak = SectionBreakKind.Continuous });

byte[] docx = report.ExportDocx();   // or report.ExportPdf();
```

---

## 8. Standalone example — `BuildContentFooter` (table + image inside a footer)

**Standalone.** A reusable branded footer: a borderless 2-column table
`[ logo | address (right-aligned) ]` followed by a centered `Page X of Y` line. The address
paragraph **wraps automatically** — no manual line-splitting — and the page number is a live
`PAGE`/`NUMPAGES` field resolved at layout (no placeholder, no "update fields" step).

> Sizing note: builder dimensions are **px** (CSS 96 dpi). 12 pt ≈ 16 px; 1 in = 96 px.

### C#

```csharp
using WordCanvas.ClearScript;
using WordCanvas.ClearScript.Builder;

public sealed record FooterInfo(string FileNumber, string Address, string CityStateZip);

// A reusable footer builder — drop it into any section's Footer(...) callback.
static void BuildContentFooter(StoryBuilder f, FooterInfo info, byte[]? logoPng)
{
    const string Font = "Times New Roman";
    const double SizePx = 16;                         // ~12pt
    var flush = new SpacingOptions { Before = 0, After = 0 };
    var borderless = new CellBorders();               // present + no edges ⇒ no lines
    var tightPad = new CellMargin(0, 2, 0, 2);

    // Leading spacer line.
    f.Paragraph(p => p.Spacing(flush));

    f.Table(t =>
    {
        t.ColFractions(0.28, 0.72);
        t.Row(r =>
        {
            // Logo cell (omitted entirely when there is no logo — see caller).
            if (logoPng is not null)
                r.Cell(c => c.Image(logoPng, "image/png",
                            new ImageOptions { WidthPx = 96, HeightPx = 40, Align = TextAlign.Left }),
                        new CellOptions { Borders = borderless, Margin = tightPad,
                                          VAlign = CellVAlign.Center,
                                          PreferredWidth = new PreferredWidth(28, WidthType.Pct) });

            // Address cell — two stacked paragraphs, right-aligned; text wraps on its own.
            r.Cell(c => c
                    .Paragraph(info.FileNumber, p => p.Align(TextAlign.Right).Font(Font).FontSize(SizePx).Spacing(flush))
                    .Paragraph($"{info.Address}, {info.CityStateZip}",
                               p => p.Align(TextAlign.Right).Font(Font).FontSize(SizePx).Spacing(flush)),
                new CellOptions { Borders = borderless, Margin = tightPad, VAlign = CellVAlign.Center,
                                  PreferredWidth = new PreferredWidth(72, WidthType.Pct) });
        });
    }, new TableOptions { WidthMode = TableWidthMode.AutofitWindow, Borders = new TableBorders() });

    // Centered "Page X of Y" with live fields.
    f.Paragraph(p => p
        .Align(TextAlign.Center).Font(Font).FontSize(SizePx)
        .Spacing(new SpacingOptions { Before = 7, After = 0 })
        .Text("Page ").PageField().Text(" of ").NumPagesField());
}

// --- using it ---
using var engine = new WordCanvasEngine();
var info = new FooterInfo("APR-2026-00042", "123 Main St", "Springfield, IL 62704");
byte[] logo = File.ReadAllBytes("logo.png");

WordDocument doc = engine.NewBuilder(new CreateOptions { PageSize = PageSizeName.Letter })
    .Footer(f => BuildContentFooter(f, info, logo))          // default band; every page
    .Paragraph("Report body…", p => p.Font("Times New Roman").FontSize(16))
    .Build();

File.WriteAllBytes("with-footer.pdf", doc.ExportPdf());
```

Per-section footers (e.g. an empty footer on the cover, the content footer everywhere else) use the
same builder callback through `SectionBreak`:

```csharp
engine.NewBuilder()
    .Footer(f => f.Paragraph())                               // cover section: empty footer
    .Paragraph("Cover")
    .SectionBreak(new SectionBreakOptions {
        BreakType = SectionBreakType.NextPage,
        Footer = f => BuildContentFooter(f, info, logo),      // body section: branded footer
    })
    .Paragraph("Body…")
    .Build();
```

### TypeScript

```ts
import { DocumentBuilder, type StoryBuilder } from "@forevka/wordcanvas/builder";
import { exportPdf } from "@forevka/wordcanvas/export";

interface FooterInfo { fileNumber: string; address: string; cityStateZip: string; }

function buildContentFooter(f: StoryBuilder, info: FooterInfo, logo?: { data: Uint8Array; mime: string }) {
  const FONT = "Times New Roman";
  const SIZE = 16;                          // px (~12pt)
  const flush = { before: 0, after: 0 } as const;
  const cell = { borders: {}, margin: { top: 0, right: 2, bottom: 0, left: 2 }, vAlign: "center" as const };

  f.paragraph((p) => p.spacing(flush));     // leading spacer

  f.table((t) => {
    t.colFractions([0.28, 0.72]);
    t.row((r) => {
      if (logo)
        r.cell((c) => c.image(logo, { widthPx: 96, heightPx: 40, align: "left" }),
                { ...cell, preferredWidth: { px: 28, type: "pct" } });

      r.cell((c) => {
        c.paragraph((p) => p.align("right").font(FONT).fontSize(SIZE).spacing(flush).text(info.fileNumber));
        c.paragraph((p) => p.align("right").font(FONT).fontSize(SIZE).spacing(flush)
                            .text(`${info.address}, ${info.cityStateZip}`));
      }, { ...cell, preferredWidth: { px: 72, type: "pct" } });
    });
  }, { widthMode: "autofitWindow", borders: {} });

  f.paragraph((p) => p.align("center").font(FONT).fontSize(SIZE).spacing({ before: 7, after: 0 })
                      .text("Page ").pageField().text(" of ").numPagesField());
}

// --- using it ---
const info: FooterInfo = { fileNumber: "APR-2026-00042", address: "123 Main St", cityStateZip: "Springfield, IL 62704" };
const logo = { data: logoBytes, mime: "image/png" };

const doc = new DocumentBuilder({ pageSize: "Letter" })
  .footer((f) => buildContentFooter(f, info, logo))
  .paragraph("Report body…", (p) => p.font("Times New Roman").fontSize(16))
  .build();

const pdf = await exportPdf(doc);
```

Key simplifications vs a raw-OpenXML footer builder: no relationship-id / `Drawing` / EMU plumbing
(`.image(bytes, …)` handles it), no manual address line-wrap heuristic (layout wraps), and no
`PAGE` field placeholder + `UpdateFields` dance (`pageField()`/`numPagesField()` resolve at layout).
```
