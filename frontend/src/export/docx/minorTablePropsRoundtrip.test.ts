// Minor & advanced table properties (issue #61) must survive a full .docx
// round-trip and be parsed from hand-written tblPr/tcPr blocks. Covers the
// table-level props (w:tblInd, w:bidiVisual, w:tblOverlap, w:tblCaption,
// w:tblDescription) and the cell-level props (w:textDirection, w:noWrap,
// w:tcFitText, w:hideMark). The importer is the oracle: a model → writeDocx →
// re-import that drops any of these would surface here.
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
const cell = (text: string, extra: Partial<TableCell> = {}): TableCell => ({ id: `c${n++}`, blocks: [cellPara(text)], ...extra });

const mkTable = (): TableBlock => ({
  kind: "table", id: "T1", revision: 0,
  colFractions: [0.5, 0.5],
  indentPx: 36,
  bidiVisual: true,
  overlap: "never",
  caption: "Leaderboard",
  description: "Alt text for the table.",
  rows: [
    { cells: [
      cell("a", { textDirection: "tbRl", noWrap: true }),
      cell("b", { fitText: true, hideMark: true }),
    ] },
  ],
});

const roundTrip = (doc: Document): Document => runImport(writeDocx(doc).bytes).doc;
const exportedXml = (doc: Document): string => strFromU8(unzipSync(writeDocx(doc).bytes)["word/document.xml"]!);
const tableOf = (d: Document): TableBlock => d.blocks.find((b): b is TableBlock => b.kind === "table")!;

// A minimal .docx whose single table carries the new tblPr + tcPr properties.
const importPropsDocx = (): Document => {
  const tblPr =
    `<w:tblPr><w:tblOverlap w:val="never"/><w:bidiVisual/>` +
    `<w:tblInd w:w="720" w:type="dxa"/><w:tblCaption w:val="Cap"/>` +
    `<w:tblDescription w:val="Desc"/></w:tblPr>`;
  const tc1 = `<w:tc><w:tcPr><w:noWrap/><w:textDirection w:val="tbRl"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>`;
  const tc2 = `<w:tc><w:tcPr><w:tcFitText/><w:hideMark/></w:tcPr><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc>`;
  const body = `<w:tbl>${tblPr}<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
    `<w:tr>${tc1}${tc2}</w:tr></w:tbl>`;
  return runImport(simpleDocx(body)).doc;
};

describe("minor & advanced table properties — issue #61", () => {
  it("parses tblPr table-level props (tblInd/bidiVisual/tblOverlap/caption/description)", () => {
    const t = tableOf(importPropsDocx());
    // 720 twips = 0.5in = 48px @96dpi.
    expect(t.indentPx).toBeCloseTo(48, 1);
    expect(t.bidiVisual).toBe(true);
    expect(t.overlap).toBe("never");
    expect(t.caption).toBe("Cap");
    expect(t.description).toBe("Desc");
  });

  it("parses tcPr cell-level props (textDirection/noWrap/fitText/hideMark)", () => {
    const t = tableOf(importPropsDocx());
    expect(t.rows[0]!.cells[0]!.textDirection).toBe("tbRl");
    expect(t.rows[0]!.cells[0]!.noWrap).toBe(true);
    expect(t.rows[0]!.cells[1]!.fitText).toBe(true);
    expect(t.rows[0]!.cells[1]!.hideMark).toBe(true);
  });

  it("emits the table-level props in CT_TblPrBase child order", () => {
    const xml = exportedXml({ section: SECTION, blocks: [mkTable()] });
    expect(xml).toContain('<w:tblOverlap w:val="never"/>');
    expect(xml).toContain("<w:bidiVisual/>");
    expect(xml).toContain('<w:tblInd w:w="540" w:type="dxa"/>');
    expect(xml).toContain('<w:tblCaption w:val="Leaderboard"/>');
    expect(xml).toContain('<w:tblDescription w:val="Alt text for the table."/>');
    // w:tblOverlap + w:bidiVisual precede w:tblW; w:tblInd precedes w:tblLayout;
    // caption/description are the final tblPr children.
    const pr = xml.slice(xml.indexOf("<w:tblPr>"), xml.indexOf("</w:tblPr>"));
    expect(pr.indexOf("<w:bidiVisual/>")).toBeLessThan(pr.indexOf("<w:tblLayout"));
    expect(pr.indexOf("<w:tblInd")).toBeLessThan(pr.indexOf("<w:tblLayout"));
    expect(pr.indexOf("<w:tblCaption")).toBeGreaterThan(pr.indexOf("<w:tblLayout"));
  });

  it("emits the cell-level props and keeps them out of an unmarked cell", () => {
    const xml = exportedXml({ section: SECTION, blocks: [mkTable()] });
    expect(xml).toContain('<w:textDirection w:val="tbRl"/>');
    expect(xml).toContain("<w:noWrap/>");
    expect(xml).toContain("<w:tcFitText/>");
    expect(xml).toContain("<w:hideMark/>");
  });

  it("preserves every property through a full export → re-import round-trip", () => {
    const out = roundTrip({ section: SECTION, blocks: [mkTable()] });
    const t = tableOf(out);
    expect(t.indentPx).toBeCloseTo(36, 0);
    expect(t.bidiVisual).toBe(true);
    expect(t.overlap).toBe("never");
    expect(t.caption).toBe("Leaderboard");
    expect(t.description).toBe("Alt text for the table.");
    expect(t.rows[0]!.cells[0]!.textDirection).toBe("tbRl");
    expect(t.rows[0]!.cells[0]!.noWrap).toBe(true);
    expect(t.rows[0]!.cells[1]!.fitText).toBe(true);
    expect(t.rows[0]!.cells[1]!.hideMark).toBe(true);
  });
});
