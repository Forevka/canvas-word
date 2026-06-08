// Table borders & shading: the OOXML cascade (table style → w:tblBorders →
// w:tcBorders, with inside/outside selection) collapsed to per-cell specs.

import { describe, expect, it } from "vitest";
import type { TableBlock } from "../../model/document";
import { runImport } from "./pipeline";
import { CONTENT_TYPES_XML, documentXml, makeDocx, REL_TYPES, relsXml, simpleDocx, stylesPartXml } from "./fixture";

const table = (r: ReturnType<typeof runImport>): TableBlock => {
  const t = r.doc.blocks.find((b) => b.kind === "table");
  expect(t).toBeDefined();
  return t as TableBlock;
};
const cell = (t: TableBlock, row: number, col: number) => t.rows[row]!.cells[col]!;

const C = (text: string, tcPr = ""): string => `<w:tc>${tcPr}<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const tbl = (tblPr: string, body: string): string => `<w:tbl><w:tblPr>${tblPr}</w:tblPr>${body}</w:tbl>`;
const grid2 = `<w:tr>${C("a")}${C("b")}</w:tr><w:tr>${C("c")}${C("d")}</w:tr>`;

describe("table borders — direct w:tblBorders", () => {
  it("leaves borders absent for an unstyled table (renderer's default grid)", () => {
    const r = runImport(simpleDocx(`<w:tbl>${grid2}</w:tbl>`));
    expect(cell(table(r), 0, 0).borders).toBeUndefined();
  });

  it("treats an explicit (empty) w:tblBorders as 'no borders', not the default grid", () => {
    const r = runImport(simpleDocx(tbl(`<w:tblBorders></w:tblBorders>`, grid2)));
    // Present-but-empty borders → defined with no edges, so the renderer draws
    // nothing instead of falling back to its gray grid (matching Word).
    expect(cell(table(r), 0, 0).borders).toEqual({});
  });

  it("treats an explicit (empty) w:tcBorders as 'no borders'", () => {
    const body = `<w:tr>${C("a", `<w:tcPr><w:tcBorders></w:tcBorders></w:tcPr>`)}${C("b")}</w:tr>`;
    expect(cell(table(runImport(simpleDocx(`<w:tbl>${body}</w:tbl>`))), 0, 0).borders).toEqual({});
  });

  it("applies outer edges from w:tblBorders only on the boundary", () => {
    const borders =
      `<w:tblBorders><w:top w:val="single" w:sz="12" w:color="FF0000"/>` +
      `<w:left w:val="single" w:sz="12" w:color="FF0000"/>` +
      `<w:bottom w:val="single" w:sz="12" w:color="FF0000"/>` +
      `<w:right w:val="single" w:sz="12" w:color="FF0000"/></w:tblBorders>`;
    const r = runImport(simpleDocx(tbl(borders, grid2)));
    const t = table(r);
    // Top-left cell: top + left are outer (set), bottom + right are interior (no
    // insideH/insideV defined → absent).
    expect(cell(t, 0, 0).borders).toEqual({
      top: { color: "#ff0000", widthPx: 2 },
      left: { color: "#ff0000", widthPx: 2 },
    });
    // Bottom-right cell: bottom + right outer.
    expect(cell(t, 1, 1).borders).toMatchObject({
      bottom: { color: "#ff0000", widthPx: 2 },
      right: { color: "#ff0000", widthPx: 2 },
    });
    expect(cell(t, 1, 1).borders!.top).toBeUndefined();
  });

  it("applies interior edges from insideH/insideV", () => {
    const borders = `<w:tblBorders><w:insideH w:val="single" w:sz="6"/><w:insideV w:val="single" w:sz="6"/></w:tblBorders>`;
    const r = runImport(simpleDocx(tbl(borders, grid2)));
    const t = table(r);
    // Top-left: bottom (interior H) and right (interior V) set; top/left absent.
    expect(cell(t, 0, 0).borders).toEqual({
      bottom: { color: "#000000", widthPx: 1 },
      right: { color: "#000000", widthPx: 1 },
    });
  });

  it("a cell w:tcBorders nil suppresses an edge the table would draw", () => {
    const borders = `<w:tblBorders><w:insideV w:val="single" w:sz="8"/></w:tblBorders>`;
    const tcNil = `<w:tcPr><w:tcBorders><w:right w:val="nil"/></w:tcBorders></w:tcPr>`;
    const body = `<w:tr>${C("a", tcNil)}${C("b")}</w:tr>`;
    const r = runImport(simpleDocx(tbl(borders, body)));
    // Cell explicitly nils its right edge → omitted despite the table's insideV.
    expect(cell(table(r), 0, 0).borders!.right).toBeUndefined();
  });

  it("maps double/dashed border styles", () => {
    const borders = `<w:tblBorders><w:top w:val="double" w:sz="8"/><w:bottom w:val="dashed" w:sz="8"/></w:tblBorders>`;
    const r = runImport(simpleDocx(tbl(borders, grid2)));
    expect(cell(table(r), 0, 0).borders!.top).toMatchObject({ style: "double" });
    expect(cell(table(r), 1, 0).borders!.bottom).toMatchObject({ style: "dashed" });
  });
});

describe("table shading — w:shd", () => {
  it("maps cell shd fill", () => {
    const body = `<w:tr>${C("a", `<w:tcPr><w:shd w:val="clear" w:fill="FFFF00"/></w:tcPr>`)}${C("b")}</w:tr>`;
    const r = runImport(simpleDocx(`<w:tbl>${body}</w:tbl>`));
    const t = table(r);
    expect(cell(t, 0, 0).shading).toBe("#ffff00");
    expect(cell(t, 0, 1).shading).toBeUndefined();
  });

  it("cell shd overrides table shd", () => {
    const tblPr = `<w:shd w:val="clear" w:fill="EEEEEE"/>`;
    const body = `<w:tr>${C("a", `<w:tcPr><w:shd w:val="clear" w:fill="FFFF00"/></w:tcPr>`)}${C("b")}</w:tr>`;
    const r = runImport(simpleDocx(tbl(tblPr, body)));
    const t = table(r);
    expect(cell(t, 0, 0).shading).toBe("#ffff00"); // cell wins
    expect(cell(t, 0, 1).shading).toBe("#eeeeee"); // table fallback
  });

  it("ignores auto/clear shading with no fill", () => {
    const body = `<w:tr>${C("a", `<w:tcPr><w:shd w:val="clear" w:fill="auto"/></w:tcPr>`)}</w:tr>`;
    const r = runImport(simpleDocx(`<w:tbl>${body}</w:tbl>`));
    expect(cell(table(r), 0, 0).shading).toBeUndefined();
  });
});

describe("table borders — table style cascade", () => {
  it("resolves borders from a referenced table style", () => {
    const styles = stylesPartXml(
      `<w:style w:type="table" w:styleId="TableGrid">
         <w:name w:val="Table Grid"/>
         <w:tblPr><w:tblBorders>
           <w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>
           <w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>
           <w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/>
         </w:tblBorders></w:tblPr>
       </w:style>`,
    );
    const bytes = makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_XML,
      "word/document.xml": documentXml(tbl(`<w:tblStyle w:val="TableGrid"/>`, grid2)),
      "word/_rels/document.xml.rels": relsXml([{ id: "rS", type: REL_TYPES.styles, target: "styles.xml" }]),
      "word/styles.xml": styles,
    });
    const r = runImport(bytes);
    const t = table(r);
    // Every cell gridded: top-left has all four (outer top/left, interior bottom/right).
    expect(cell(t, 0, 0).borders).toMatchObject({
      top: { widthPx: round(4 / 6) },
      left: {},
      bottom: {},
      right: {},
    });
    // Direct table borders override the style (none here), cell overrides table.
    expect(cell(t, 1, 1).borders!.bottom).toBeDefined();
  });
});

const round = (v: number): number => Math.round(v * 100) / 100;

describe("table cell margins — w:tcMar / w:tblCellMar cascade", () => {
  const mar = (sides: Record<string, number>): string =>
    `<w:tcMar>${Object.entries(sides).map(([s, w]) => `<w:${s} w:w="${w}" w:type="dxa"/>`).join("")}</w:tcMar>`;

  it("falls back to Word defaults (0 vertical, ~7.2px horizontal) when nothing is set", () => {
    const r = runImport(simpleDocx(`<w:tbl>${grid2}</w:tbl>`));
    expect(cell(table(r), 0, 0).margin).toEqual({ top: 0, right: 7.2, bottom: 0, left: 7.2 });
  });

  it("resolves a cell's w:tcMar per side, defaulting the sides it omits", () => {
    // The report's tables specify only top/bottom = 40 twips (≈2.67px).
    const body = `<w:tr>${C("a", `<w:tcPr>${mar({ top: 40, bottom: 40 })}</w:tcPr>`)}${C("b")}</w:tr>`;
    const r = runImport(simpleDocx(`<w:tbl>${body}</w:tbl>`));
    expect(cell(table(r), 0, 0).margin).toEqual({ top: 2.67, right: 7.2, bottom: 2.67, left: 7.2 });
  });

  it("cascades cell w:tcMar over table w:tblCellMar over the default, per side", () => {
    const tblPr = `<w:tblCellMar><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="200" w:type="dxa"/><w:right w:w="200" w:type="dxa"/></w:tblCellMar>`;
    const body = `<w:tr>${C("a", `<w:tcPr>${mar({ top: 40 })}</w:tcPr>`)}${C("b")}</w:tr>`;
    const r = runImport(simpleDocx(tbl(tblPr, body)));
    const t = table(r);
    // Cell overrides only top (40→2.67); bottom/left/right inherit the table's tblCellMar.
    expect(cell(t, 0, 0).margin).toEqual({ top: 2.67, right: 13.33, bottom: 8, left: 13.33 });
    // The sibling cell with no tcMar takes the table's tblCellMar on every side.
    expect(cell(t, 0, 1).margin).toEqual({ top: 8, right: 13.33, bottom: 8, left: 13.33 });
  });
});
