// Table row properties (w:trPr) must survive a full .docx round-trip and parse
// from a hand-written trPr block. The importer is the oracle: a model → writeDocx
// → re-import that drops trHeight / cantSplit / tblHeader surfaces here.
import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import type { CharStyle, Document, Paragraph, SectionProps, TableBlock, TableRow } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import { simpleDocx } from "../../import/docx/fixture";
import { writeDocx } from "./writeDocx";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 14, bold: false, italic: false, underline: false, strikethrough: false, color: "#202124" };
const PARA = { align: "left" as const, lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

let n = 0;
const cellPara = (text: string): Paragraph => ({ kind: "paragraph", id: `p${n++}`, revision: 0, runs: [{ text, style: { ...CHAR } }], style: { ...PARA } });
const cell = (text: string) => ({ id: `c${n++}`, blocks: [cellPara(text)] });
const row = (a: string, b: string, props?: TableRow["props"]): TableRow => ({ cells: [cell(a), cell(b)], ...(props ? { props } : {}) });

const mkTable = (): TableBlock => ({
  kind: "table", id: "T1", revision: 0,
  colFractions: [0.5, 0.5],
  rows: [
    row("Header A", "Header B", { repeatHeader: true }),
    row("exact", "row", { height: { value: 48, rule: "exact" } }),
    row("atleast", "row", { height: { value: 60, rule: "atLeast" } }),
    row("kept", "whole", { cantSplit: true }),
    row("plain", "row"),
  ],
});

const roundTrip = (doc: Document): Document => runImport(writeDocx(doc).bytes).doc;
const exportedXml = (doc: Document): string => strFromU8(unzipSync(writeDocx(doc).bytes)["word/document.xml"]!);
const tableOf = (d: Document): TableBlock => d.blocks.find((b): b is TableBlock => b.kind === "table")!;

// A minimal .docx whose table rows carry hand-written w:trPr blocks.
const trWith = (trPr: string): string =>
  `<w:tr>${trPr}<w:tc><w:tcPr/><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:tcPr/><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr>`;
const importTrPrDocx = (): Document => {
  const body = `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
    trWith(`<w:trPr><w:tblHeader/></w:trPr>`) +
    trWith(`<w:trPr><w:trHeight w:val="720" w:hRule="exact"/></w:trPr>`) +
    trWith(`<w:trPr><w:trHeight w:val="900" w:hRule="atLeast"/></w:trPr>`) +
    trWith(`<w:trPr><w:cantSplit/></w:trPr>`) +
    trWith(`<w:trPr><w:trHeight w:val="500" w:hRule="auto"/></w:trPr>`) + // auto → dropped
    `</w:tbl>`;
  return runImport(simpleDocx(body)).doc;
};

describe("table row properties — w:trPr", () => {
  it("parses w:trHeight (atLeast/exact), w:cantSplit and w:tblHeader from trPr", () => {
    const t = tableOf(importTrPrDocx());
    expect(t.rows[0]!.props?.repeatHeader).toBe(true);
    expect(t.rows[1]!.props?.height).toEqual({ value: 48, rule: "exact" }); // 720 twips = 48px
    expect(t.rows[2]!.props?.height).toEqual({ value: 60, rule: "atLeast" }); // 900 twips = 60px
    expect(t.rows[3]!.props?.cantSplit).toBe(true);
    // hRule="auto" is a pure hint — dropped, leaving no row props.
    expect(t.rows[4]!.props).toBeUndefined();
  });

  it("emits w:trPr elements only for rows that set a property", () => {
    const xml = exportedXml({ section: SECTION, blocks: [mkTable()] });
    expect(xml).toContain("<w:tblHeader/>");
    expect(xml).toContain('<w:trHeight w:val="720" w:hRule="exact"/>'); // 48px
    expect(xml).toContain('<w:trHeight w:val="900" w:hRule="atLeast"/>'); // 60px
    expect(xml).toContain("<w:cantSplit/>");
    // The plain trailing row emits a bare <w:tr> (no <w:trPr>) before its first cell.
    expect(xml).toContain("<w:tr><w:tc>");
  });

  it("preserves every row property through a full export → re-import round-trip", () => {
    const out = roundTrip({ section: SECTION, blocks: [mkTable()] });
    const t = tableOf(out);
    expect(t.rows[0]!.props?.repeatHeader).toBe(true);
    expect(t.rows[1]!.props?.height).toEqual({ value: 48, rule: "exact" });
    expect(t.rows[2]!.props?.height).toEqual({ value: 60, rule: "atLeast" });
    expect(t.rows[3]!.props?.cantSplit).toBe(true);
    expect(t.rows[4]!.props).toBeUndefined();
  });
});
