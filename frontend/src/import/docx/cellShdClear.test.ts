// Ground-truth import + round-trip test for issue #150 (follow-up to #147): a
// CELL's w:shd is tri-state just like a paragraph's — absent (inherit through
// cell → table w:tblPr/w:shd → table style), present-with-fill (use it), and
// present-but-EMPTY (`<w:shd w:val="clear" w:fill="auto"/>` / `w:val="nil"` —
// Word's "Shading → No Color", an explicit clear of an inherited fill). The
// importer used to collapse the third state into "inherit", so a cleared cell
// re-shaded from the table / table style. These assertions are against the
// SOURCE XML's meaning (a round-trip can't catch it — the wrong model is a
// stable fixed point) plus an export re-emission check.

import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import type { Document, TableBlock } from "@cw/shared";
import { writeDocx } from "../../export/docx/writeDocx";
import { CONTENT_TYPES_XML, documentXml, makeDocx, REL_TYPES, relsXml, simpleDocx, stylesPartXml } from "./fixture";
import { runImport } from "./pipeline";

const table = (d: Document): TableBlock => {
  const t = d.blocks.find((b) => b.kind === "table");
  expect(t).toBeDefined();
  return t as TableBlock;
};
const cell = (t: TableBlock, row: number, col: number) => t.rows[row]!.cells[col]!;
const C = (text: string, tcPr = ""): string => `<w:tc>${tcPr}<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const CLEAR = `<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:tcPr>`;
const NIL = `<w:tcPr><w:shd w:val="nil"/></w:tcPr>`;
const exportedXml = (d: Document): string => strFromU8(unzipSync(writeDocx(d).bytes)["word/document.xml"]!);

// A table whose w:tblPr/w:shd carries a fill; first cell clears it, second inherits.
const tableLevelDoc = (clearTcPr: string) =>
  runImport(
    simpleDocx(
      `<w:tbl><w:tblPr><w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/></w:tblPr>` +
        `<w:tr>${C("cleared", clearTcPr)}${C("inherits")}</w:tr></w:tbl>`,
    ),
  ).doc;

// A table referencing a table style whose wholeTable shading carries a fill.
const STYLES = stylesPartXml(
  `<w:style w:type="table" w:styleId="TblGreen">
     <w:name w:val="Table Green"/>
     <w:tblPr><w:shd w:val="clear" w:color="auto" w:fill="D9EAD3"/></w:tblPr>
   </w:style>`,
);
const styleLevelDoc = (clearTcPr: string) =>
  runImport(
    makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_XML,
      "word/document.xml": documentXml(
        `<w:tbl><w:tblPr><w:tblStyle w:val="TblGreen"/></w:tblPr>` +
          `<w:tr>${C("cleared", clearTcPr)}${C("inherits")}</w:tr></w:tbl>`,
      ),
      "word/_rels/document.xml.rels": relsXml([{ id: "rS", type: REL_TYPES.styles, target: "styles.xml" }]),
      "word/styles.xml": STYLES,
    }),
  ).doc;

describe("cell w:shd tri-state (issue #150)", () => {
  it("inherits the table-level w:shd fill when the cell has no w:shd (sanity)", () => {
    const t = table(tableLevelDoc(""));
    expect(cell(t, 0, 0).shading).toBe("#eeeeee");
    expect(cell(t, 0, 1).shading).toBe("#eeeeee");
  });

  it('clears the table-level fill on an explicit "No Color" (fill="auto")', () => {
    const t = table(tableLevelDoc(CLEAR));
    expect(cell(t, 0, 0).shading).toBeUndefined();
    expect(cell(t, 0, 0).shadingCleared).toBe(true);
    expect(cell(t, 0, 1).shading).toBe("#eeeeee"); // sibling still inherits
  });

  it("clears the table-level fill on a nil w:shd", () => {
    const t = table(tableLevelDoc(NIL));
    expect(cell(t, 0, 0).shading).toBeUndefined();
    expect(cell(t, 0, 0).shadingCleared).toBe(true);
  });

  it("inherits the table STYLE's wholeTable fill when the cell has no w:shd (sanity)", () => {
    const t = table(styleLevelDoc(""));
    expect(cell(t, 0, 0).shading).toBe("#d9ead3");
    expect(cell(t, 0, 1).shading).toBe("#d9ead3");
  });

  it("clears the table STYLE's fill on an explicit No Color", () => {
    const t = table(styleLevelDoc(CLEAR));
    expect(cell(t, 0, 0).shading).toBeUndefined();
    expect(cell(t, 0, 0).shadingCleared).toBe(true);
    expect(cell(t, 0, 1).shading).toBe("#d9ead3"); // sibling still inherits
  });
});

describe("cell w:shd explicit-clear export/round-trip (issue #150)", () => {
  it("re-emits an empty w:shd for a cleared cell under a shaded table", () => {
    const xml = exportedXml(tableLevelDoc(CLEAR));
    expect(xml).toContain(`<w:shd w:val="clear" w:color="auto" w:fill="auto"/>`);
  });

  it("survives import → export → import (the clear is not lost)", () => {
    const back = runImport(writeDocx(tableLevelDoc(CLEAR)).bytes).doc;
    const t = table(back);
    expect(cell(t, 0, 0).shading).toBeUndefined();
    expect(cell(t, 0, 0).shadingCleared).toBe(true);
    expect(cell(t, 0, 1).shading).toBe("#eeeeee");
  });
});
