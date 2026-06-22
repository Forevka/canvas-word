// Nested content controls must survive a full .docx round-trip: export
// reconstructs nested w:sdt (inline via run sdtPath, block-level via Block.sdtPath)
// by longest-common-prefix grouping; re-import rebuilds the same ancestry.

import { describe, expect, it } from "vitest";
import type { CharStyle, Document, FieldDef, Paragraph, SdtProps, SectionProps, TableBlock } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import { writeDocx } from "./writeDocx";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA = { align: "left" as const, lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };
const rt = (tag: string): SdtProps => ({ type: "richText", tag });

const para = (id: string, runs: { text: string; path?: string[] }[], sdtPath?: string[]): Paragraph => ({
  kind: "paragraph", id, revision: 0, style: PARA,
  runs: runs.map((r) => ({ text: r.text, style: r.path ? { ...CHAR, sdtPath: r.path } : { ...CHAR } })),
  ...(sdtPath ? { sdtPath } : {}),
});

// section(s0) wraps: a paragraph with a nested inline control, and a table whose
// value cell holds its own inline control.
const table: TableBlock = {
  kind: "table", id: "t1", revision: 0, sdtPath: ["s0"],
  rows: [{ cells: [
    { id: "c1", blocks: [para("cp1", [{ text: "Label" }])] },
    { id: "c2", blocks: [para("cp2", [{ text: "200", path: ["val"] }])] },
  ] }],
};

const doc: Document = {
  section: SECTION,
  sdts: { s0: rt("section"), inl: rt("inner"), val: rt("value") },
  blocks: [
    // Block-level section (s0) on the paragraph; the inline control (inl) only on
    // run "b". A run's full chain is block.sdtPath ++ run path.
    para("p1", [
      { text: "a" },
      { text: "b", path: ["inl"] },
      { text: "c" },
    ], ["s0"]),
    table,
  ],
};

const tagsOf = (d: Document, path: string[] | undefined): string[] => (path ?? []).map((id) => d.sdts![id]!.tag!);

describe("nested SDT .docx round-trip", () => {
  it("preserves nested inline + block-level controls and does not fragment the section", () => {
    const back = runImport(writeDocx(doc).bytes).doc;

    // Exactly three controls survive (no fragmentation of the section wrapper).
    expect(Object.keys(back.sdts!).length).toBe(3);

    const p = back.blocks.find((b): b is Paragraph => b.kind === "paragraph")!;
    const chain = (t: string): string[] => [...tagsOf(back, p.sdtPath), ...tagsOf(back, p.runs.find((r) => r.text === t)!.style.sdtPath)];
    expect(chain("a")).toEqual(["section"]);
    expect(chain("b")).toEqual(["section", "inner"]); // nesting preserved
    expect(chain("c")).toEqual(["section"]);

    // The table is wrapped by the section; the value cell keeps its own control.
    const tbl = back.blocks.find((b): b is TableBlock => b.kind === "table")!;
    expect(tagsOf(back, tbl.sdtPath)).toEqual(["section"]);
    const valuePara = tbl.rows[0]!.cells[1]!.blocks[0] as Paragraph;
    expect(valuePara.sdtPath).toBeUndefined(); // not re-wrapped by the section
    expect(tagsOf(back, valuePara.runs.find((r) => r.text === "200")!.style.sdtPath)).toEqual(["value"]);
  });

  it("preserves an inline FIELD nested inside nested content controls", () => {
    // A DATE field whose result run sits inside an inner control, inside an outer
    // one: sdtPath (control ancestry) and fieldId (field membership) are orthogonal
    // markers on the same run and must both survive the round-trip.
    const dateDef: FieldDef = { id: "f1", instruction: ' DATE \\@ "M/d/yyyy" ', name: "DATE", kind: "builtin", spec: { type: "DATE", format: "M/d/yyyy" } };
    const fieldDoc: Document = {
      section: SECTION,
      sdts: { o: rt("outer"), i: rt("inner") },
      fields: { f1: dateDef },
      blocks: [
        {
          kind: "paragraph", id: "p1", revision: 0, style: PARA,
          runs: [
            { text: "x", style: { ...CHAR, sdtPath: ["o"] } },
            { text: "1/1/2026", style: { ...CHAR, sdtPath: ["o", "i"], fieldId: "f1" } },
            { text: "y", style: { ...CHAR, sdtPath: ["o"] } },
          ],
        },
      ],
    };
    const back = runImport(writeDocx(fieldDoc).bytes).doc;

    // The field definition survives.
    const date = Object.values(back.fields ?? {}).find((d) => d.name === "DATE");
    expect(date?.spec).toEqual({ type: "DATE", format: "M/d/yyyy" });

    // The result run carries BOTH the nested control ancestry and the field id.
    const p = back.blocks[0] as Paragraph;
    const fieldRun = p.runs.find((r) => r.text === "1/1/2026")!;
    expect(tagsOf(back, fieldRun.style.sdtPath)).toEqual(["outer", "inner"]);
    expect(fieldRun.style.fieldId).toBe(date!.id);
    // The controls themselves still round-trip (two, un-fragmented).
    expect(Object.keys(back.sdts!).length).toBe(2);
  });
});
