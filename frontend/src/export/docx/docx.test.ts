// DOCX export tested by ROUND-TRIP: import a .docx -> model A, write it back out,
// re-import -> model B, and assert the model survived. The existing importer is
// the oracle, so anything the writer mis-encodes shows up as drift.

import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import type { Block, Document, Paragraph, TableBlock } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import {
  CONTENT_TYPES_XML,
  documentXml,
  makeDocx,
  relsXml,
  REL_TYPES,
  simpleDocx,
} from "../../import/docx/fixture";
import { writeDocx } from "./writeDocx";

const roundTrip = (doc: Document): Document => runImport(writeDocx(doc).bytes).doc;
const exportedDocumentXml = (doc: Document): string =>
  strFromU8(unzipSync(writeDocx(doc).bytes)["word/document.xml"]!);
const paras = (d: Document): Paragraph[] => d.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const text = (p: Paragraph): string => p.runs.map((r) => r.text).join("");
const tables = (d: Document): TableBlock[] => d.blocks.filter((b): b is TableBlock => b.kind === "table");
const cellText = (b: Block): string =>
  b.kind === "paragraph" ? text(b) : "";

describe("DOCX export — round trip", () => {
  it("preserves paragraph text and run formatting", () => {
    const a = runImport(
      simpleDocx(
        `<w:p><w:r><w:rPr><w:b/><w:i/><w:color w:val="FF0000"/><w:sz w:val="48"/><w:rFonts w:ascii="Arial"/></w:rPr><w:t>Styled</w:t></w:r>` +
          `<w:r><w:t xml:space="preserve"> plain</w:t></w:r></w:p>`,
      ),
    ).doc;
    const b = roundTrip(a);
    expect(text(paras(b)[0]!)).toBe("Styled plain");
    const styled = paras(b)[0]!.runs[0]!.style;
    expect(styled.bold).toBe(true);
    expect(styled.italic).toBe(true);
    expect(styled.color).toBe("#ff0000");
    expect(styled.fontSizePx).toBe(32); // 48 half-pt = 24pt = 32px
    expect(styled.fontFamily).toBe("Arial, serif");
  });

  it("preserves alignment, spacing and indents", () => {
    const a = runImport(
      simpleDocx(
        `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/>` +
          `<w:ind w:left="720" w:right="480" w:firstLine="360"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
      ),
    ).doc;
    const s = paras(roundTrip(a))[0]!.style;
    expect(s.align).toBe("center");
    expect(s.spaceBeforePx).toBeCloseTo(16, 1);
    expect(s.spaceAfterPx).toBeCloseTo(8, 1);
    expect(s.lineHeight).toBeCloseTo(1.5, 2);
    expect(s.indentLeftPx).toBeCloseTo(48, 1);
    expect(s.indentRightPx).toBeCloseTo(32, 1); // 480 twips = 32px
    expect(s.indentFirstLinePx).toBeCloseTo(24, 1);
  });

  it("round-trips paragraph borders (w:pBdr) and paragraph shading (w:shd)", () => {
    const pBdr =
      `<w:pBdr>` +
      `<w:top w:val="single" w:sz="8" w:color="1A73E8"/>` +
      `<w:left w:val="single" w:sz="8" w:color="1A73E8"/>` +
      `<w:bottom w:val="double" w:sz="12" w:color="1A73E8"/>` +
      `<w:right w:val="dashed" w:sz="8" w:color="1A73E8"/>` +
      `</w:pBdr>`;
    const a = runImport(
      simpleDocx(`<w:p><w:pPr>${pBdr}<w:shd w:val="clear" w:color="auto" w:fill="EEF4FF"/></w:pPr><w:r><w:t>boxed</w:t></w:r></w:p>`),
    ).doc;
    const sa = paras(a)[0]!.style;
    expect(sa.shading).toBe("#eef4ff");
    expect(sa.borders?.top?.color).toBe("#1a73e8");
    expect(sa.borders?.bottom?.style).toBe("double");
    expect(sa.borders?.right?.style).toBe("dashed");

    const sb = paras(roundTrip(a))[0]!.style;
    expect(sb.shading).toBe("#eef4ff");
    expect(sb.borders?.top?.color).toBe("#1a73e8");
    expect(sb.borders?.top?.widthPx).toBeCloseTo(sa.borders!.top!.widthPx, 5);
    expect(sb.borders?.bottom?.style).toBe("double");
    expect(sb.borders?.right?.style).toBe("dashed");
    // The left edge survives as a plain single border (no `style` key on a single edge).
    expect(sb.borders?.left).toBeDefined();
    expect(sb.borders?.left?.style).toBeUndefined();
    expect(sb.borders?.left?.color).toBe("#1a73e8");
  });

  it("round-trips fixed line spacing (w:lineRule exact and atLeast)", () => {
    const a = runImport(
      simpleDocx(
        `<w:p><w:pPr><w:spacing w:line="420" w:lineRule="exact"/></w:pPr><w:r><w:t>exact</w:t></w:r></w:p>` +
          `<w:p><w:pPr><w:spacing w:line="360" w:lineRule="atLeast"/></w:pPr><w:r><w:t>atLeast</w:t></w:r></w:p>`,
      ),
    ).doc;
    const b = roundTrip(a);
    const [exact, atLeast] = paras(b);
    expect(exact!.style.lineRule).toBe("exact");
    expect(exact!.style.lineHeightPx).toBeCloseTo(28, 1); // 420 twips = 28px
    expect(atLeast!.style.lineRule).toBe("atLeast");
    expect(atLeast!.style.lineHeightPx).toBeCloseTo(24, 1); // 360 twips = 24px
    // And it survives in the serialized XML as the right rule, not "auto".
    const xml = exportedDocumentXml(b);
    expect(xml).toContain(`w:line="420"`);
    expect(xml).toContain(`w:lineRule="exact"`);
    expect(xml).toContain(`w:lineRule="atLeast"`);
  });

  it("round-trips table cell margins (w:tcMar) — top/bottom must survive", () => {
    // Regression: the writer used to drop cell margins, so top/bottom re-imported as
    // Word's default 0 — shrinking every row and drifting page count by ~5%.
    const a = runImport(
      simpleDocx(
        `<w:tbl><w:tblPr><w:tblCellMar>` +
          `<w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/>` +
          `<w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
          `<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>` +
          `<w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
      ),
    ).doc;
    const ma = tables(a)[0]!.rows[0]!.cells[0]!.margin!;
    expect(ma.top).toBeCloseTo(6.67, 1); // 100 twips
    const mb = tables(roundTrip(a))[0]!.rows[0]!.cells[0]!.margin!;
    expect(mb.top).toBeCloseTo(6.67, 1);
    expect(mb.bottom).toBeCloseTo(6.67, 1);
    expect(mb.left).toBeCloseTo(7.2, 1);
  });

  it("re-emits table-level borders / shading / cell-margins at tblPr level (issue #48)", () => {
    // w:tblBorders (incl. the interior insideH/insideV edges), w:tblPr/w:shd and
    // w:tblCellMar are parsed as table-WIDE defaults. They must survive a
    // Word→edit→Word cycle at the table level, not only as the resolved per-cell
    // copies — otherwise the table-level defaults are lost and every cell is bloated.
    const tblBorders =
      `<w:tblBorders>` +
      `<w:top w:val="single" w:sz="12" w:color="1A73E8"/>` +
      `<w:left w:val="single" w:sz="12" w:color="1A73E8"/>` +
      `<w:bottom w:val="single" w:sz="12" w:color="1A73E8"/>` +
      `<w:right w:val="single" w:sz="12" w:color="1A73E8"/>` +
      `<w:insideH w:val="dashed" w:sz="6" w:color="34A853"/>` +
      `<w:insideV w:val="dashed" w:sz="6" w:color="34A853"/></w:tblBorders>`;
    const cellMar =
      `<w:tblCellMar><w:top w:w="120" w:type="dxa"/><w:left w:w="150" w:type="dxa"/>` +
      `<w:bottom w:w="120" w:type="dxa"/><w:right w:w="150" w:type="dxa"/></w:tblCellMar>`;
    const body =
      `<w:tbl><w:tblPr>${tblBorders}<w:shd w:val="clear" w:color="auto" w:fill="EEF5FF"/>${cellMar}</w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
      `<w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr>` +
      `<w:tr><w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>d</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const a = runImport(simpleDocx(body)).doc;
    const ta = tables(a)[0]!;
    // Import captures the table-level defaults on the model.
    expect(ta.defaultBorders?.top?.color).toBe("#1a73e8");
    expect(ta.defaultBorders?.insideH?.style).toBe("dashed");
    expect(ta.defaultBorders?.insideV?.color).toBe("#34a853");
    expect(ta.defaultShading).toBe("#eef5ff");
    expect(ta.defaultCellMargin?.left).toBeCloseTo(10, 1); // 150 twips = 10px
    expect(ta.defaultCellMargin?.top).toBeCloseTo(8, 1); // 120 twips = 8px

    // The exported tblPr carries the three table-level elements (not only per-cell).
    const xml = exportedDocumentXml(a);
    const tblPr = xml.slice(xml.indexOf("<w:tblPr>"), xml.indexOf("</w:tblPr>"));
    expect(tblPr).toContain("<w:tblBorders>");
    expect(tblPr).toContain("w:insideH");
    expect(tblPr).toContain("w:insideV");
    expect(tblPr).toContain(`w:fill="eef5ff"`);
    expect(tblPr).toContain("<w:tblCellMar>");

    // …and they survive the full round-trip on the model.
    const tb = tables(roundTrip(a))[0]!;
    expect(tb.defaultBorders?.top?.color).toBe("#1a73e8");
    expect(tb.defaultBorders?.insideV?.color).toBe("#34a853");
    expect(tb.defaultBorders?.insideH?.style).toBe("dashed");
    expect(tb.defaultShading).toBe("#eef5ff");
    expect(tb.defaultCellMargin?.top).toBeCloseTo(8, 1);
    expect(tb.defaultCellMargin?.left).toBeCloseTo(10, 1);
  });

  it("preserves a table with gridSpan, vMerge, borders and shading", () => {
    const border = `<w:tcBorders><w:top w:val="single" w:sz="12" w:color="0000FF"/><w:bottom w:val="single" w:sz="12" w:color="0000FF"/><w:left w:val="single" w:sz="12" w:color="0000FF"/><w:right w:val="single" w:sz="12" w:color="0000FF"/></w:tcBorders>`;
    const body =
      `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:shd w:fill="FFFF00"/></w:tcPr><w:p><w:r><w:t>span2</w:t></w:r></w:p></w:tc></w:tr>` +
      `<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/>${border}</w:tcPr><w:p><w:r><w:t>owner</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:p><w:r><w:t>r2c2</w:t></w:r></w:p></w:tc></w:tr>` +
      `<w:tr><w:tc><w:tcPr><w:vMerge w:val="continue"/></w:tcPr><w:p/></w:tc>` +
      `<w:tc><w:p><w:r><w:t>r3c2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const a = runImport(simpleDocx(body)).doc;
    const b = roundTrip(a);
    const ta = tables(a)[0]!;
    const tb = tables(b)[0]!;
    expect(tb.rows.length).toBe(ta.rows.length);
    // row 0: a single cell spanning 2 columns, shaded yellow
    expect(tb.rows[0]!.cells[0]!.colSpan).toBe(2);
    expect(tb.rows[0]!.cells[0]!.shading).toBe("#ffff00");
    // row 1 col 1: rowSpan owner with blue borders
    const owner = tb.rows[1]!.cells[0]!;
    expect(owner.rowSpan).toBe(2);
    expect(owner.borders?.top?.color).toBe("#0000ff");
    expect(cellText(owner.blocks[0]!)).toBe("owner");
    // the continue row drops its merged cell (one cell present)
    expect(tb.rows[2]!.cells.length).toBe(1);
  });

  it("preserves an explicit no-borders table (does not revert to the default grid)", () => {
    // An author's explicit "no borders" — empty <w:tblBorders/> — imports to a
    // present-but-empty CellBorders {} so the renderer suppresses its gray
    // default grid. The export must keep that intent or the table reopens gray.
    const body =
      `<w:tbl><w:tblPr><w:tblBorders/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
      `<w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const a = runImport(simpleDocx(body)).doc;
    const cellA = tables(a)[0]!.rows[0]!.cells[0]!;
    expect(cellA.borders).toBeDefined(); // imported as "borders specified, none drawn"
    // The explicit (empty) w:tblBorders container is kept at table level too, so the
    // "no borders" declaration round-trips at tblPr — not just per-cell (issue #48).
    expect(tables(a)[0]!.defaultBorders).toEqual({});
    const b = roundTrip(a);
    const cellB = tables(b)[0]!.rows[0]!.cells[0]!;
    // Must stay defined (empty) — undefined would draw the gray default grid.
    expect(cellB.borders).toBeDefined();
    expect(cellB.borders!.top).toBeUndefined();
    expect(tables(b)[0]!.defaultBorders).toEqual({});
  });

  it("mirrors a vertical merge's borders onto its continue cells (Word draws the merged box)", () => {
    // A rowSpan owner carries the merged region's full borders, but Word renders
    // a merged cell's sides/bottom from its CONTINUE cells, not the restart cell.
    // A 2x2 table, fully bordered, with col 0 merged down both rows.
    const border = `<w:tblBorders><w:top w:val="single" w:sz="8" w:color="000000"/><w:left w:val="single" w:sz="8" w:color="000000"/><w:bottom w:val="single" w:sz="8" w:color="000000"/><w:right w:val="single" w:sz="8" w:color="000000"/><w:insideH w:val="single" w:sz="8" w:color="000000"/><w:insideV w:val="single" w:sz="8" w:color="000000"/></w:tblBorders>`;
    const body =
      `<w:tbl><w:tblPr>${border}</w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>m</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr>` +
      `<w:tr><w:tc><w:tcPr><w:vMerge w:val="continue"/></w:tcPr><w:p/></w:tc>` +
      `<w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const a = runImport(simpleDocx(body)).doc;
    const xml = exportedDocumentXml(a);
    // Isolate the synthesized continue cell.
    const cont = xml.slice(xml.indexOf('<w:vMerge w:val="continue"/>'));
    const tcBorders = cont.slice(cont.indexOf("<w:tcBorders>"), cont.indexOf("</w:tcBorders>"));
    expect(tcBorders).toContain("<w:left"); // side borders repeat on the band
    expect(tcBorders).toContain("<w:right");
    expect(tcBorders).toContain("<w:bottom"); // final band owns the region's bottom
    expect(tcBorders).not.toContain("<w:top"); // interior to the merge
  });

  it("preserves a combined colSpan+rowSpan merge without inflating rowSpan", () => {
    // One cell spanning 2 cols AND 2 rows. The continue band must round-trip as a
    // single gridSpan'd cell — one continue PER COLUMN would make the importer
    // bump the owner's rowSpan once per column (2 -> 3), drifting the model.
    const body =
      `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>big</w:t></w:r></w:p></w:tc></w:tr>` +
      `<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge w:val="continue"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`;
    const a = runImport(simpleDocx(body)).doc;
    const owner = tables(a)[0]!.rows[0]!.cells[0]!;
    expect(owner.colSpan).toBe(2);
    expect(owner.rowSpan).toBe(2);
    const b = roundTrip(a);
    const ownerB = tables(b)[0]!.rows[0]!.cells[0]!;
    expect(ownerB.colSpan).toBe(2);
    expect(ownerB.rowSpan).toBe(2); // not 3 — the continue band stayed one cell
    expect(tables(b)[0]!.rows.length).toBe(2);
  });

  it("preserves list membership and marker format", () => {
    const numbering = `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>` +
      `<w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const docx = makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_XML,
      "word/document.xml": documentXml(`<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p>`),
      "word/_rels/document.xml.rels": relsXml([{ id: "rId1", type: REL_TYPES.numbering, target: "numbering.xml" }]),
      "word/numbering.xml": numbering,
    });
    const a = runImport(docx).doc;
    const b = roundTrip(a);
    const p = paras(b)[0]!;
    expect(p.style.list).toBeDefined();
    const listId = p.style.list!.listId;
    expect(b.lists?.[listId]?.levels[0]?.format).toBe("decimal");
  });

  it("preserves a list on a paragraph INSIDE a table cell", () => {
    const numbering = `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>` +
      `<w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const docx = makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_XML,
      "word/document.xml": documentXml(
        `<w:tbl><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>` +
          `<w:tr><w:tc><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>cell item</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
      ),
      "word/_rels/document.xml.rels": relsXml([{ id: "rId1", type: REL_TYPES.numbering, target: "numbering.xml" }]),
      "word/numbering.xml": numbering,
    });
    const a = runImport(docx).doc;
    const cellParaA = tables(a)[0]!.rows[0]!.cells[0]!.blocks[0]!;
    expect(cellParaA.kind === "paragraph" && cellParaA.style.list).toBeDefined(); // imported onto the cell paragraph
    const b = roundTrip(a);
    const cellParaB = tables(b)[0]!.rows[0]!.cells[0]!.blocks[0]!;
    expect(cellParaB.kind === "paragraph" && cellParaB.style.list).toBeDefined(); // and survives export → re-import
    const listId = cellParaB.kind === "paragraph" ? cellParaB.style.list!.listId : "";
    expect(b.lists?.[listId]?.levels[0]?.format).toBe("decimal");
    expect(cellText(cellParaB)).toBe("cell item");
  });

  it("preserves per-section page geometry", () => {
    // A4 (11906 x 16838 twips) body section.
    const a = runImport(
      simpleDocx(`<w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`),
    ).doc;
    const b = roundTrip(a);
    expect(b.section.pageWidthPx).toBeCloseTo(a.section.pageWidthPx, 0);
    expect(b.section.pageHeightPx).toBeCloseTo(a.section.pageHeightPx, 0);
  });

  it("round-trips header/footer band distances (w:header/w:footer not clobbered to 720)", () => {
    // Regression: the writer hardcoded w:header/w:footer="720", dropping real distances.
    const a = runImport(
      simpleDocx(
        `<w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
          `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="1080" w:footer="900"/></w:sectPr>`,
      ),
    ).doc;
    expect(a.section.headerDistancePx).toBeCloseTo(72, 0); // 1080 twips
    expect(a.section.footerDistancePx).toBeCloseTo(60, 0); // 900 twips
    const b = roundTrip(a);
    expect(b.section.headerDistancePx).toBeCloseTo(72, 0);
    expect(b.section.footerDistancePx).toBeCloseTo(60, 0);
  });

  it("round-trips landscape orientation (emits w:orient, keeps w > h)", () => {
    const a = runImport(
      simpleDocx(
        `<w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/></w:sectPr>`,
      ),
    ).doc;
    expect(a.section.pageWidthPx).toBeGreaterThan(a.section.pageHeightPx);
    expect(exportedDocumentXml(a)).toContain('w:orient="landscape"');
    const b = roundTrip(a);
    expect(b.section.pageWidthPx).toBeCloseTo(a.section.pageWidthPx, 0);
    expect(b.section.pageHeightPx).toBeCloseTo(a.section.pageHeightPx, 0);
    expect(b.section.pageWidthPx).toBeGreaterThan(b.section.pageHeightPx);
  });

  it("round-trips column separator + per-column widths (w:sep, w:col)", () => {
    const a = runImport(
      simpleDocx(
        `<w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr>` +
          `<w:cols w:num="2" w:sep="1" w:equalWidth="0"><w:col w:w="3000" w:space="240"/><w:col w:w="4000"/></w:cols></w:sectPr>`,
      ),
    ).doc;
    expect(a.section.columns?.sep).toBe(true);
    expect(a.section.columns?.cols).toHaveLength(2);
    const b = roundTrip(a);
    expect(b.section.columns?.count).toBe(2);
    expect(b.section.columns?.sep).toBe(true);
    expect(b.section.columns?.cols).toHaveLength(2);
    expect(b.section.columns?.cols?.[0]?.widthPx).toBeCloseTo(200, 0); // 3000 twips
    expect(b.section.columns?.cols?.[1]?.widthPx).toBeCloseTo(266.67, 0); // 4000 twips
  });

  it("round-trips page color (w:background)", () => {
    const a = runImport(simpleDocx(`<w:p><w:r><w:t>x</w:t></w:r></w:p>`)).doc;
    a.section.pageColorHex = "#ffcc00";
    expect(exportedDocumentXml(a)).toContain("w:background");
    const b = roundTrip(a);
    expect(b.section.pageColorHex?.toLowerCase()).toBe("#ffcc00");
  });

  it("round-trips page borders (w:pgBorders)", () => {
    const a = runImport(simpleDocx(`<w:p><w:r><w:t>x</w:t></w:r></w:p>`)).doc;
    const edge = { style: "single" as const, widthPx: 2, color: "#ff0000" };
    a.section.pageBorders = {
      offsetFrom: "page",
      top: { ...edge }, right: { ...edge }, bottom: { ...edge }, left: { ...edge },
    };
    const b = roundTrip(a);
    expect(b.section.pageBorders?.offsetFrom).toBe("page");
    expect(b.section.pageBorders?.top?.style).toBe("single");
    expect(b.section.pageBorders?.top?.color.toLowerCase()).toBe("#ff0000");
    expect(b.section.pageBorders?.top?.widthPx).toBeCloseTo(2, 1);
  });

  it("preserves external and in-document hyperlinks", () => {
    const docx = makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_XML,
      "word/document.xml": documentXml(
        `<w:p><w:hyperlink r:id="rId9"><w:r><w:t>ext</w:t></w:r></w:hyperlink></w:p>` +
          `<w:p><w:hyperlink w:anchor="mark"><w:r><w:t>jump</w:t></w:r></w:hyperlink></w:p>`,
      ),
      "word/_rels/document.xml.rels": relsXml([
        { id: "rId9", type: REL_TYPES.hyperlink, target: "https://example.com/", external: true },
      ]),
    });
    const b = roundTrip(runImport(docx).doc);
    expect(paras(b)[0]!.runs[0]!.style.link).toBe("https://example.com/");
    expect(paras(b)[1]!.runs[0]!.style.link).toBe("#mark");
  });

  it("preserves footnotes", () => {
    const footnotes =
      `<?xml version="1.0"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:footnote w:id="2"><w:p><w:r><w:t>the note</w:t></w:r></w:p></w:footnote></w:footnotes>`;
    const docx = makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_XML,
      "word/document.xml": documentXml(`<w:p><w:r><w:t>body</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>`),
      "word/_rels/document.xml.rels": relsXml([{ id: "rId1", type: REL_TYPES.footnotes, target: "footnotes.xml" }]),
      "word/footnotes.xml": footnotes,
    });
    const a = runImport(docx).doc;
    const b = roundTrip(a);
    const refRun = paras(b)[0]!.runs.find((r) => r.style.footnoteRef);
    expect(refRun).toBeDefined();
    const noteId = refRun!.style.footnoteRef!;
    expect(b.footnotes?.[noteId]?.[0] && text(b.footnotes[noteId][0] as Paragraph)).toBe("the note");
  });

  it("preserves endnotes", () => {
    const endnotes =
      `<?xml version="1.0"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:endnote w:id="2"><w:p><w:r><w:t>the endnote</w:t></w:r></w:p></w:endnote></w:endnotes>`;
    const docx = makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_XML,
      "word/document.xml": documentXml(`<w:p><w:r><w:t>body</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p>`),
      "word/_rels/document.xml.rels": relsXml([{ id: "rId1", type: REL_TYPES.endnotes, target: "endnotes.xml" }]),
      "word/endnotes.xml": endnotes,
    });
    const a = runImport(docx).doc;
    const b = roundTrip(a);
    const refRun = paras(b)[0]!.runs.find((r) => r.style.endnoteRef);
    expect(refRun).toBeDefined();
    const noteId = refRun!.style.endnoteRef!;
    expect(b.endnotes?.[noteId]?.[0] && text(b.endnotes[noteId][0] as Paragraph)).toBe("the endnote");
  });

  it("preserves an embedded image through a supplied bytes map", () => {
    // 1x1 PNG.
    const png = Uint8Array.from(
      atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="),
      (c) => c.charCodeAt(0),
    );
    const doc: Document = {
      section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
      blocks: [
        { kind: "image", id: "img1", revision: 0, src: "blob:fake", widthPx: 96, heightPx: 96, align: "center" },
      ],
    };
    const { bytes, warnings } = writeDocx(doc, { "blob:fake": png });
    expect(warnings.find((w) => w.code === "image-unresolved")).toBeUndefined();
    const b = runImport(bytes).doc;
    const img = b.blocks.find((bl) => bl.kind === "image");
    expect(img).toBeDefined();
    expect((img as Extract<Block, { kind: "image" }>).widthPx).toBeCloseTo(96, 0);
  });

  it("round-trips a behind-text anchored (wrapNone) background image as wp:anchor", () => {
    const png = Uint8Array.from(
      atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="),
      (c) => c.charCodeAt(0),
    );
    const char = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
    const paraStyle = { align: "left" as const, lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
    const doc: Document = {
      section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
      blocks: [
        {
          kind: "image", id: "bg", revision: 0, src: "blob:fake", widthPx: 816, heightPx: 1056, align: "left",
          anchor: { behind: true, offsetXPx: -98, offsetYPx: -238, relFromH: "margin", relFromV: "paragraph", decorative: true, z: 42 },
        },
        { kind: "paragraph", id: "p1", revision: 0, runs: [{ text: "PARTY INVITATION", style: char }], style: paraStyle },
      ],
    };
    const { bytes } = writeDocx(doc, { "blob:fake": png });
    // Exported as a floating anchor, not inline.
    const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
    expect(xml).toContain("wp:anchor");
    expect(xml).toContain('behindDoc="1"');
    expect(xml).toContain("wp:wrapNone");
    expect(xml).not.toContain("wp:inline");
    // Re-import keeps the anchor + the body text.
    const b = runImport(bytes).doc;
    const img = b.blocks.find((bl) => bl.kind === "image") as Extract<Block, { kind: "image" }>;
    expect(img.anchor).toBeDefined();
    expect(img.anchor!.behind).toBe(true);
    expect(img.anchor!.offsetXPx).toBeCloseTo(-98, 0);
    expect(img.anchor!.offsetYPx).toBeCloseTo(-238, 0);
    expect(img.anchor!.relFromH).toBe("margin");
    expect(img.anchor!.relFromV).toBe("paragraph");
    expect(img.anchor!.z).toBe(42); // relativeHeight round-trips the stacking order
    expect(img.wrap).toBeUndefined();
    expect(b.blocks.some((bl) => bl.kind === "paragraph")).toBe(true);
  });

  it("preserves bookmarks", () => {
    const docx = makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_XML,
      "word/document.xml": documentXml(
        `<w:bookmarkStart w:id="1" w:name="target"/><w:p><w:r><w:t>anchored</w:t></w:r></w:p><w:bookmarkEnd w:id="1"/>`,
      ),
    });
    const a = runImport(docx).doc;
    const b = roundTrip(a);
    expect(Object.keys(b.bookmarks ?? {})).toContain("target");
  });

  it("round-trips a bookmark's character range (start and end offsets)", () => {
    const a = runImport(
      simpleDocx(`<w:p><w:bookmarkStart w:id="3" w:name="bm"/><w:r><w:t>hello world</w:t></w:r><w:bookmarkEnd w:id="3"/></w:p>`),
    ).doc;
    const ra = a.bookmarks!["bm"]!;
    expect(ra.start.offset).toBe(0);
    expect(ra.end.offset).toBe(11); // after "hello world"
    expect(ra.start.blockId).toBe(paras(a)[0]!.id);
    const b = roundTrip(a);
    const rb = b.bookmarks!["bm"]!;
    expect(rb.start.offset).toBe(0);
    expect(rb.end.offset).toBe(11);
    expect(rb.start.blockId).toBe(paras(b)[0]!.id);
  });

  it("round-trips an imported TOC field as a LIVE field with the preserved instruction", () => {
    // A real Word TOC: outer TOC field wrapping HYPERLINK+PAGEREF entries, with the
    // headings carrying _Toc bookmarks. Import flattens+marks the entries; export
    // must re-emit a live TOC field with the SAME instruction (not a hardcoded one).
    const entry = (anchor: string, label: string, page: string): string =>
      `<w:hyperlink w:anchor="${anchor}"><w:r><w:t>${label}</w:t></w:r><w:r><w:tab/></w:r>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF ${anchor} \\h </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${page}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink>`;
    const head = (anchor: string, label: string): string => {
      const id = anchor.replace(/\D/g, "");
      return `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="${id}" w:name="${anchor}"/>` +
        `<w:r><w:t>${label}</w:t></w:r><w:bookmarkEnd w:id="${id}"/></w:p>`;
    };
    const body =
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> TOC \\o "1-5" \\h </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${entry("_Toc1", "Chapter One", "3")}</w:p>` +
      `<w:p>${entry("_Toc2", "Chapter Two", "5")}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>` +
      head("_Toc1", "Chapter One") + head("_Toc2", "Chapter Two");

    const a = runImport(simpleDocx(body)).doc;
    expect(a.tocInstruction).toContain('TOC \\o "1-5"'); // captured verbatim
    expect(a.blocks.filter((b) => b.kind === "paragraph" && b.style.tocEntry).length).toBe(2); // marked

    const xml = exportedDocumentXml(a);
    expect(xml).toContain('TOC \\o "1-5"'); // preserved instruction re-emitted (not default \z \u)
    expect(xml).toContain("PAGEREF _Toc1"); // live field, not flattened text
    expect(xml).toContain('w:anchor="_Toc1"'); // \h hyperlink
  });

  it("tags a TOC entry's runs with its target anchor (nested HYPERLINK + PAGEREF)", () => {
    // Word writes a TOC entry as a HYPERLINK field wrapping the label and a nested
    // PAGEREF field for the page number — neither is a w:hyperlink element. The
    // importer must still surface the "_Toc1" target as an in-document link so
    // "recalculate TOC" can map the entry to its heading.
    const entry =
      `<w:p>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> HYPERLINK \\l "_Toc1" </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:r><w:t>Chapter One</w:t></w:r><w:r><w:tab/></w:r>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:r><w:t>7</w:t></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
      `</w:p>`;
    const a = runImport(simpleDocx(entry)).doc;
    const p = paras(a)[0]!;
    expect(text(p)).toBe("Chapter One\t7");
    expect(p.runs.some((r) => r.style.link === "#_Toc1")).toBe(true);
    // Survives a round-trip (links re-emit as w:hyperlink w:anchor).
    const b = roundTrip(a);
    expect(paras(b)[0]!.runs.some((r) => r.style.link === "#_Toc1")).toBe(true);
  });

  it("emits footer {page}/{pages} as live PAGE/NUMPAGES fields (body tokens stay literal)", () => {
    const CHAR = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
    const PARA = { align: "left" as const, lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
    const doc: Document = {
      section: {
        pageWidthPx: 816,
        pageHeightPx: 1056,
        marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
        footer: [{ kind: "paragraph", id: "f1", revision: 0, runs: [{ text: "Page {page} of {pages}", style: CHAR }], style: PARA }],
      },
      blocks: [{ kind: "paragraph", id: "p1", revision: 0, runs: [{ text: "body {page}", style: CHAR }], style: PARA }],
    };
    const zip = unzipSync(writeDocx(doc).bytes);

    // Footer tokens -> live fields.
    const footerName = Object.keys(zip).find((n) => /^word\/footer\d+\.xml$/.test(n));
    expect(footerName).toBeDefined();
    const footerXml = strFromU8(zip[footerName!]!);
    expect(footerXml).toContain("PAGE \\* MERGEFORMAT");
    expect(footerXml).toContain("NUMPAGES \\* MERGEFORMAT");

    // Body tokens stay literal text (symmetric with import, which only tokenizes bands).
    const bodyXml = strFromU8(zip["word/document.xml"]!);
    expect(bodyXml).not.toContain("w:fldSimple");
    expect(bodyXml).toContain("body {page}");

    // The footer re-imports with the tokens intact, so layout still substitutes per page.
    const b = runImport(writeDocx(doc).bytes).doc;
    const footerText = (b.section.footer ?? []).map((bl) => (bl.kind === "paragraph" ? text(bl) : "")).join("");
    expect(footerText).toBe("Page {page} of {pages}");
  });

  it("round-trips hidden text (w:vanish) instead of dropping it", () => {
    const a = runImport(
      simpleDocx(
        `<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>secret</w:t></w:r><w:r><w:t>shown</w:t></w:r></w:p>`,
      ),
    ).doc;
    // exporter must emit w:vanish for the hidden run
    expect(exportedDocumentXml(a)).toContain("<w:vanish");
    const b = roundTrip(a);
    const p = paras(b)[0]!;
    const secret = p.runs.find((r) => r.text === "secret")!;
    expect(secret.style.hidden).toBe(true);
    expect(p.runs.find((r) => r.text === "shown")!.style.hidden).toBeUndefined();
  });

  it("produces an archive the importer reads without error", () => {
    const a = runImport(simpleDocx(`<w:p><w:r><w:t>hello</w:t></w:r></w:p>`)).doc;
    const { bytes, warnings } = writeDocx(a);
    // PK zip magic
    expect(bytes[0]).toBe(0x50);
    expect(bytes[1]).toBe(0x4b);
    expect(warnings).toEqual([]);
    expect(paras(runImport(bytes).doc).length).toBeGreaterThan(0);
  });
});
