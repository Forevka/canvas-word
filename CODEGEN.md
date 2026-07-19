# Reverse Builder — `.docx` → DocumentBuilder code

`@forevka/wordcanvas/codegen` is the **inverse** of [`./builder`](./BUILDER.md).
It takes a `.docx` (or an in-memory `Document`) and emits editable TypeScript
that calls the fluent `DocumentBuilder` API to reconstruct it — so a document
authored **visually** in Word (or the editor) becomes a **code template** a
developer maintains, parameterizes, and re-runs.

```ts
import { docxToBuilderCode } from "@forevka/wordcanvas/codegen";

const { code, uncovered, importWarnings } = docxToBuilderCode(docxBytes);
// `code` is a ready-to-save .ts module exporting buildDocument(): Document
// `uncovered` lists any model fields the fluent API cannot yet express
```

The emitted module looks like hand-written builder code:

```ts
import { DocumentBuilder } from "@forevka/wordcanvas/builder";
import type { Document } from "@forevka/wordcanvas/builder";

export function buildDocument(): Document {
  const b = DocumentBuilder.create();
  b.pageSetup({ pageSize: { pageWidthPx: 794, pageHeightPx: 1123 }, margins: { top: 96, right: 72, bottom: 96, left: 72 } });
  b.paragraph().withStyle("Title").text("Q3 Financials");
  b.paragraph().withStyle("Heading1").text("Summary");
  b.paragraph().text("Revenue grew ").text("18%", { bold: true, color: "#0b7a0b" }).text(" year over year.");
  b.table((t) => {
    t.row((r) => {
      r.cell((c) => { c.paragraph().text("Metric", { bold: true }); });
      r.cell((c) => { c.paragraph().text("Q3", { bold: true }); });
    });
  });
  return b.build();
}
```

The workflow it enables: **design in Word → convert to a template → wire in your
data.** Replace the literal strings with your data model and the generated
function becomes a report/letter/contract generator.

## API

```ts
// Pure model → source (isomorphic, no import needed).
emitBuilderCode(doc: Document, opts?: CodegenOptions): CodegenResult;

// Import .docx bytes, then emit.
docxToBuilderCode(docx: ArrayBuffer | Uint8Array, opts?: CodegenOptions): DocxCodegenResult;

interface CodegenOptions {
  functionName?: string;  // default "buildDocument"
  importFrom?: string;    // default "@forevka/wordcanvas/builder"
}
interface CodegenResult { code: string; uncovered: UncoveredField[]; }
interface DocxCodegenResult extends CodegenResult { importWarnings: string[]; }
```

In the editor, develop mode (`develop: true`) adds an **Export builder code**
button to the Developer ribbon tab that generates and downloads `template.ts`
for the current document.

## Fidelity

Runs reproduce **exactly** through the generic `.text(text, patch)` escape hatch:
the emitter passes the full structural delta of each run's `CharStyle` against the
resolved baseline the builder itself would produce (the same `resolveStyle`
cascade `BuilderContext` uses), so every character field — including the
self-contained `equation`/`symbol`/`charStyleId` — comes back verbatim.
Paragraph properties map to their fluent methods the same way, and the
stylesheet / list definitions / table styles reproduce verbatim via
`create({ stylesheet })` / `.listDefinition` / `.tableStyle`. The round-trip test
suite (`frontend/src/codegen/codegen.test.ts`) **executes** the generated source
and asserts the reconstructed model equals the original.

## The coverage report — deliberately not lossy-silent

The fluent API is not 100% surjective onto the model. Where a field cannot be
expressed, the emitter **records it in `uncovered`** (node path + field + a
sample value + a note) instead of dropping it silently. That report is the
backlog for **growing the builder to cover the frequently-used patches** — the
first such gap, per-paragraph list membership, was closed this way by adding
`ParagraphBuilder.listItem(listId, level)`.

Currently reported (reconstruction not yet automated): inline/region **fields**,
**content controls** (SDT), **footnotes/endnotes**, **bookmarks**, the **TOC**
field, **block equations** (AST→MathML serializer not wired), **custom blocks**,
Word identity metadata (`paraId`/`textId`/`docId`), page `pageColorHex`/
`pageBorders`, and a few advanced table/section properties. Each surfaces in
`uncovered` with a note on what would close it.

It lives at `frontend/src/codegen/` (published as the `./codegen` subpath),
alongside the import pipeline it reuses.
