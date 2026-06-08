// DOCX export tested by ROUND-TRIP: import a .docx -> model A, write it back out,
// re-import -> model B, and assert the model survived. The existing importer is
// the oracle, so anything the writer mis-encodes shows up as drift.

import { describe, expect, it } from "vitest";
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

  it("preserves per-section page geometry", () => {
    // A4 (11906 x 16838 twips) body section.
    const a = runImport(
      simpleDocx(`<w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`),
    ).doc;
    const b = roundTrip(a);
    expect(b.section.pageWidthPx).toBeCloseTo(a.section.pageWidthPx, 0);
    expect(b.section.pageHeightPx).toBeCloseTo(a.section.pageHeightPx, 0);
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
