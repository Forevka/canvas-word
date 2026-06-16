// Inline field objects must survive a full .docx round-trip: export wraps a
// fieldId-tagged run run in a complex field (begin/instrText/separate/result/end);
// re-import reconstructs the FieldDef + re-tags the result run.
import { describe, expect, it } from "vitest";
import type { CharStyle, Document, FieldDef, Paragraph, SectionProps } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import { writeDocx } from "./writeDocx";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA = { align: "left" as const, lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

const pageDef: FieldDef = { id: "f1", instruction: " PAGE \\* MERGEFORMAT ", name: "PAGE", kind: "builtin", spec: { type: "PAGE" } };
const dateDef: FieldDef = { id: "f2", instruction: ' DATE \\@ "M/d/yyyy" ', name: "DATE", kind: "builtin", spec: { type: "DATE", format: "M/d/yyyy" } };

const doc: Document = {
  section: SECTION,
  fields: { f1: pageDef, f2: dateDef },
  blocks: [
    {
      kind: "paragraph", id: "p1", revision: 0, style: PARA,
      runs: [
        { text: "Page ", style: CHAR },
        { text: "{page}", style: { ...CHAR, fieldId: "f1" } },
        { text: " · ", style: CHAR },
        { text: "1/1/2026", style: { ...CHAR, fieldId: "f2" } },
      ],
    },
  ],
};

describe("inline field .docx round-trip", () => {
  it("reconstructs PAGE (token) and DATE (materialized) field objects after export→import", () => {
    const bytes = writeDocx(doc).bytes;
    const back = runImport(bytes).doc;

    const defs = Object.values(back.fields ?? {});
    const page = defs.find((d) => d.name === "PAGE");
    const date = defs.find((d) => d.name === "DATE");
    expect(page?.kind).toBe("builtin");
    expect(page?.spec).toEqual({ type: "PAGE" });
    expect(date?.spec).toEqual({ type: "DATE", format: "M/d/yyyy" });

    const p = back.blocks[0] as Paragraph;
    const pageRun = p.runs.find((r) => r.text === "{page}");
    const dateRun = p.runs.find((r) => r.text === "1/1/2026");
    expect(pageRun?.style.fieldId).toBe(page!.id); // PAGE result re-tagged
    expect(dateRun?.style.fieldId).toBe(date!.id); // DATE cached result re-tagged
  });
});
