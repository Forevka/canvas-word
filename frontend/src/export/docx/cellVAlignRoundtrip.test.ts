// Cell vertical alignment (w:tcPr/w:vAlign) must survive a full .docx round-trip
// and be parsed from a hand-written tcPr block. The importer is the oracle: a model
// → writeDocx → re-import that drops vAlign would surface here.
import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import type { CharStyle, Document, Paragraph, SectionProps, TableBlock, TableCell } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import { simpleDocx } from "../../import/docx/fixture";
import { writeDocx } from "./writeDocx";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 14, bold: false, italic: false, underline: false, strikethrough: false, color: "#202124" };
const PARA = { align: "left" as const, lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

let n = 0;
const cellPara = (text: string): Paragraph => ({ kind: "paragraph", id: `p${n++}`, revision: 0, runs: [{ text, style: { ...CHAR } }], style: { ...PARA } });
const cell = (text: string, vAlign?: TableCell["vAlign"]): TableCell => ({ id: `c${n++}`, blocks: [cellPara(text)], ...(vAlign ? { vAlign } : {}) });
const mkTable = (): TableBlock => ({
  kind: "table", id: "T1", revision: 0,
  colFractions: [0.34, 0.33, 0.33],
  rows: [
    { cells: [cell("top", "top"), cell("center", "center"), cell("bottom", "bottom")] },
    { cells: [cell("plain"), cell("c", "center"), cell("b", "bottom")] },
  ],
});

const roundTrip = (doc: Document): Document => runImport(writeDocx(doc).bytes).doc;
const exportedXml = (doc: Document): string => strFromU8(unzipSync(writeDocx(doc).bytes)["word/document.xml"]!);
const tableOf = (d: Document): TableBlock => d.blocks.find((b): b is TableBlock => b.kind === "table")!;

// A minimal .docx whose single table carries a w:vAlign on each first-row cell.
const tcWithVAlign = (text: string, val?: string): string =>
  `<w:tc><w:tcPr>${val ? `<w:vAlign w:val="${val}"/>` : ""}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const importVAlignDocx = (): Document => {
  const body = `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
    `<w:tr>${tcWithVAlign("a", "center")}${tcWithVAlign("b", "bottom")}</w:tr></w:tbl>`;
  return runImport(simpleDocx(body)).doc;
};

describe("cell vertical alignment — w:vAlign", () => {
  it("parses w:vAlign from tcPr into the cell model (center/bottom; top/absent omitted)", () => {
    const t = tableOf(importVAlignDocx());
    expect(t.rows[0]!.cells[0]!.vAlign).toBe("center");
    expect(t.rows[0]!.cells[1]!.vAlign).toBe("bottom");
  });

  it("emits w:vAlign only for center/bottom (top stays the implicit default)", () => {
    const xml = exportedXml({ section: SECTION, blocks: [mkTable()] });
    expect(xml).toContain('<w:vAlign w:val="center"/>');
    expect(xml).toContain('<w:vAlign w:val="bottom"/>');
    // The explicit "top" cell and the plain cell must NOT emit a w:vAlign.
    expect(xml).not.toContain('<w:vAlign w:val="top"/>');
  });

  it("preserves center/bottom through a full export → re-import round-trip", () => {
    const out = roundTrip({ section: SECTION, blocks: [mkTable()] });
    const t = tableOf(out);
    // Row 0: explicit top normalizes to absent; center/bottom survive.
    expect(t.rows[0]!.cells[0]!.vAlign).toBeUndefined();
    expect(t.rows[0]!.cells[1]!.vAlign).toBe("center");
    expect(t.rows[0]!.cells[2]!.vAlign).toBe("bottom");
    // Row 1: plain cell stays absent; center/bottom survive.
    expect(t.rows[1]!.cells[0]!.vAlign).toBeUndefined();
    expect(t.rows[1]!.cells[1]!.vAlign).toBe("center");
    expect(t.rows[1]!.cells[2]!.vAlign).toBe("bottom");
  });
});
