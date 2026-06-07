// End-to-end pipeline tests: synthesized .docx bytes → model Document.
// These are the golden tests from IMPORT.md — every lossy policy is asserted
// alongside the warning it must emit.

import { readFileSync } from "node:fs"; // typed by node-fs.d.ts — no @types/node in this DOM-only project
import { describe, expect, it } from "vitest";
import type { Paragraph, TableBlock } from "../../model/document";
import { runImport } from "./pipeline";
import { ImportError } from "./types";
import { CONTENT_TYPES_XML, documentXml, makeDocx, simpleDocx } from "./fixture";

const importBody = (body: string) => runImport(simpleDocx(body));

const para = (b: unknown): Paragraph => {
  expect((b as Paragraph).kind).toBe("paragraph");
  return b as Paragraph;
};

const warningCodes = (r: ReturnType<typeof runImport>): string[] => r.warnings.map((w) => w.code);

describe("docx pipeline — paragraphs and runs", () => {
  it("imports paragraphs with direct run formatting", () => {
    const r = importBody(
      `<w:p><w:r><w:t>Plain </w:t></w:r><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>bold-italic</w:t></w:r></w:p>` +
        `<w:p><w:r><w:rPr><w:color w:val="FF0000"/><w:sz w:val="48"/><w:rFonts w:ascii="Arial"/></w:rPr><w:t>big red</w:t></w:r></w:p>`,
    );
    expect(r.doc.blocks).toHaveLength(2);

    const p1 = para(r.doc.blocks[0]);
    expect(p1.runs).toHaveLength(2);
    expect(p1.runs[0]).toMatchObject({ text: "Plain ", style: { bold: false, italic: false } });
    expect(p1.runs[1]).toMatchObject({ text: "bold-italic", style: { bold: true, italic: true } });

    const p2 = para(r.doc.blocks[1]);
    expect(p2.runs[0]!.style).toMatchObject({ color: "#ff0000", fontSizePx: 32, fontFamily: "Arial, serif" });
    expect(r.warnings).toHaveLength(0);
  });

  it("honors explicit off-toggles (w:b w:val=0)", () => {
    const r = importBody(`<w:p><w:r><w:rPr><w:b w:val="0"/><w:u w:val="none"/></w:rPr><w:t>x</w:t></w:r></w:p>`);
    expect(para(r.doc.blocks[0]).runs[0]!.style).toMatchObject({ bold: false, underline: false });
  });

  it("merges adjacent same-styled runs (normalization invariant)", () => {
    const r = importBody(`<w:p><w:r><w:t>one </w:t></w:r><w:r><w:t>two</w:t></w:r></w:p>`);
    const p = para(r.doc.blocks[0]);
    expect(p.runs).toHaveLength(1);
    expect(p.runs[0]!.text).toBe("one two");
  });

  it("preserves significant whitespace in w:t", () => {
    const r = importBody(
      `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>a</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>b</w:t></w:r></w:p>`,
    );
    const p = para(r.doc.blocks[0]);
    expect(p.runs.map((x) => x.text).join("")).toBe("a b");
  });

  it("decodes XML entities", () => {
    const r = importBody(`<w:p><w:r><w:t>Q &amp; A &lt;tags&gt;</w:t></w:r></w:p>`);
    expect(para(r.doc.blocks[0]).runs[0]!.text).toBe("Q & A <tags>");
  });

  it("maps paragraph properties: alignment, spacing, indents", () => {
    const r = importBody(
      `<w:p><w:pPr><w:jc w:val="both"/><w:spacing w:before="240" w:after="120" w:line="360"/>` +
        `<w:ind w:left="720" w:hanging="360"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
    );
    expect(para(r.doc.blocks[0]).style).toMatchObject({
      align: "justify",
      spaceBeforePx: 16,
      spaceAfterPx: 8,
      lineHeight: 1.5,
      indentLeftPx: 48,
      indentFirstLinePx: -24, // hanging
    });
  });

  it("records w:pStyle as namedStyle for the future StyleResolver", () => {
    const r = importBody(`<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>title</w:t></w:r></w:p>`);
    expect(para(r.doc.blocks[0]).style.namedStyle).toBe("Heading1");
  });
});

describe("docx pipeline — lossy policies emit warnings", () => {
  it("converts tabs to fixed spaces", () => {
    const r = importBody(`<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>`);
    expect(para(r.doc.blocks[0]).runs[0]!.text).toBe("a    b");
    expect(warningCodes(r)).toContain("tabs");
  });

  it("splits soft line breaks into paragraphs with hugged spacing", () => {
    const r = importBody(
      `<w:p><w:pPr><w:spacing w:before="240" w:after="240"/></w:pPr>` +
        `<w:r><w:t>line1</w:t><w:br/><w:t>line2</w:t></w:r></w:p>`,
    );
    expect(r.doc.blocks).toHaveLength(2);
    const [p1, p2] = [para(r.doc.blocks[0]), para(r.doc.blocks[1])];
    expect(p1.runs[0]!.text).toBe("line1");
    expect(p2.runs[0]!.text).toBe("line2");
    expect(p1.style.spaceAfterPx).toBe(0); // split halves hug each other
    expect(p2.style.spaceBeforePx).toBe(0);
    expect(p2.style.spaceAfterPx).toBe(16); // outer spacing survives on the last
    expect(warningCodes(r)).toContain("soft-breaks");
  });

  it("preserves empty lines from consecutive and trailing soft breaks", () => {
    const r = importBody(`<w:p><w:r><w:t>a</w:t><w:br/><w:br/><w:t>b</w:t><w:br/></w:r></w:p>`);
    const texts = r.doc.blocks.map((b) => para(b).runs.map((x) => x.text).join(""));
    expect(texts).toEqual(["a", "", "b", ""]);
  });

  it("maps w:pageBreakBefore onto ParaStyle.pageBreakBefore", () => {
    const r = importBody(
      `<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>new page</w:t></w:r></w:p>`,
    );
    expect(para(r.doc.blocks[0]).style.pageBreakBefore).toBe(true);
  });

  it("turns an inline page break into pageBreakBefore on the following content", () => {
    const r = importBody(`<w:p><w:r><w:t>page1</w:t><w:br w:type="page"/><w:t>page2</w:t></w:r></w:p>`);
    expect(r.doc.blocks).toHaveLength(2);
    expect(para(r.doc.blocks[0]).style.pageBreakBefore).toBeUndefined();
    expect(para(r.doc.blocks[1]).style.pageBreakBefore).toBe(true);
    expect(warningCodes(r)).not.toContain("soft-breaks"); // faithful, not lossy
  });

  it("carries a trailing page break on an empty follower paragraph", () => {
    const r = importBody(`<w:p><w:r><w:t>end of page</w:t><w:br w:type="page"/></w:r></w:p>`);
    const last = para(r.doc.blocks[1]);
    expect(last.runs[0]!.text).toBe("");
    expect(last.style.pageBreakBefore).toBe(true);
  });

  it("treats a mid-document section break as a page boundary", () => {
    const r = importBody(
      `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p>` +
        `<w:p><w:r><w:t>second section</w:t></w:r></w:p>`,
    );
    expect(para(r.doc.blocks[1]).style.pageBreakBefore).toBe(true);
    expect(warningCodes(r)).toContain("multiple-sections"); // geometry is still last-wins
  });

  it("does NOT page-break on continuous section breaks", () => {
    const r = importBody(
      `<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/></w:sectPr></w:pPr></w:p>` +
        `<w:p><w:r><w:t>same page</w:t></w:r></w:p>`,
    );
    expect(para(r.doc.blocks[1]).style.pageBreakBefore).toBeUndefined();
  });

  it("unwraps content controls (block and inline)", () => {
    const r = importBody(
      `<w:sdt><w:sdtContent><w:p><w:r><w:t>in control</w:t></w:r></w:p></w:sdtContent></w:sdt>` +
        `<w:p><w:sdt><w:sdtContent><w:r><w:t>inline control</w:t></w:r></w:sdtContent></w:sdt></w:p>`,
    );
    expect(para(r.doc.blocks[0]).runs[0]!.text).toBe("in control");
    expect(para(r.doc.blocks[1]).runs[0]!.text).toBe("inline control");
    expect(warningCodes(r)).toContain("sdt-unwrapped");
  });

  it("flattens hyperlinks to styled text", () => {
    const r = importBody(
      `<w:p><w:hyperlink r:id="rId4" xmlns:r="x"><w:r><w:rPr><w:u/></w:rPr><w:t>click me</w:t></w:r></w:hyperlink></w:p>`,
    );
    const run = para(r.doc.blocks[0]).runs[0]!;
    expect(run.text).toBe("click me");
    expect(run.style.underline).toBe(true);
    expect(warningCodes(r)).toContain("hyperlinks-flattened");
  });

  it("keeps cached field results, drops instructions", () => {
    const r = importBody(
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>` +
        `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r>` +
        `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
    );
    expect(para(r.doc.blocks[0]).runs[0]!.text).toBe("7");
  });

  it("drops hidden text (w:vanish)", () => {
    const r = importBody(
      `<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>secret</w:t></w:r><w:r><w:t>visible</w:t></w:r></w:p>`,
    );
    const p = para(r.doc.blocks[0]);
    expect(p.runs.map((x) => x.text).join("")).toBe("visible");
    expect(warningCodes(r)).toContain("hidden-text");
  });

});

describe("docx pipeline — tables", () => {
  const cell = (text: string, tcPr = ""): string =>
    `<w:tc>${tcPr}<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;

  it("imports a uniform table", () => {
    const r = importBody(
      `<w:tbl><w:tr>${cell("a")}${cell("b")}</w:tr><w:tr>${cell("c")}${cell("d")}</w:tr></w:tbl>`,
    );
    const tbl = r.doc.blocks[0] as TableBlock;
    expect(tbl.kind).toBe("table");
    expect(tbl.rows).toHaveLength(2);
    expect(tbl.rows[0]!.cells).toHaveLength(2);
    expect(para(tbl.rows[1]!.cells[1]!.blocks[0]).runs[0]!.text).toBe("d");
    expect(r.warnings).toHaveLength(0);
  });

  it("maps gridSpan to colSpan (horizontal merge preserved)", () => {
    const spanned = cell("wide", `<w:tcPr><w:gridSpan w:val="2"/></w:tcPr>`);
    const r = importBody(`<w:tbl><w:tr>${spanned}</w:tr><w:tr>${cell("a")}${cell("b")}</w:tr></w:tbl>`);
    const tbl = r.doc.blocks[0] as TableBlock;
    expect(tbl.rows[0]!.cells).toHaveLength(1);
    expect(tbl.rows[0]!.cells[0]!.colSpan).toBe(2);
    expect(tbl.rows[1]!.cells).toHaveLength(2);
    expect(r.warnings).toHaveLength(0); // faithful — no longer a lossy mapping
  });

  it("splits vertical merges with a warning (no row spans in the model)", () => {
    const restart = cell("tall", `<w:tcPr><w:vMerge w:val="restart"/></w:tcPr>`);
    const cont = `<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>`;
    const r = importBody(
      `<w:tbl><w:tr>${restart}${cell("b")}</w:tr><w:tr>${cont}${cell("d")}</w:tr></w:tbl>`,
    );
    expect(warningCodes(r)).toContain("cell-vmerge");
  });

  it("derives colFractions from w:tblGrid", () => {
    const r = importBody(
      `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="1000"/></w:tblGrid>` +
        `<w:tr>${cell("a")}${cell("b")}</w:tr></w:tbl>`,
    );
    expect((r.doc.blocks[0] as TableBlock).colFractions).toEqual([0.75, 0.25]);
  });

  it("preserves nested tables one level deep", () => {
    const inner = `<w:tbl><w:tr>${cell("inner")}</w:tr></w:tbl>`;
    const r = importBody(
      `<w:tbl><w:tr><w:tc><w:p><w:r><w:t>outer</w:t></w:r></w:p>${inner}</w:tc></w:tr></w:tbl>`,
    );
    const outer = r.doc.blocks[0] as TableBlock;
    const cellBlocks = outer.rows[0]!.cells[0]!.blocks;
    expect(cellBlocks[0]!.kind).toBe("paragraph");
    expect(cellBlocks[1]!.kind).toBe("table");
    const nested = cellBlocks[1] as TableBlock;
    expect(para(nested.rows[0]!.cells[0]!.blocks[0]).runs[0]!.text).toBe("inner");
  });

  it("every cell keeps at least one editable paragraph", () => {
    const r = importBody(`<w:tbl><w:tr><w:tc></w:tc></w:tr></w:tbl>`);
    const tbl = r.doc.blocks[0] as TableBlock;
    const blocks = tbl.rows[0]!.cells[0]!.blocks;
    expect(blocks).toHaveLength(1);
    expect(para(blocks[0]).runs[0]!.text).toBe("");
  });
});

describe("docx pipeline — section and container", () => {
  it("maps sectPr page geometry to px", () => {
    const r = importBody(
      `<w:p><w:r><w:t>x</w:t></w:r></w:p>` +
        `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`,
    );
    expect(r.doc.section).toMatchObject({
      pageWidthPx: 816,
      pageHeightPx: 1056,
      marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
    });
  });

  it("defaults the section when sectPr is absent", () => {
    const r = importBody(`<w:p><w:r><w:t>x</w:t></w:r></w:p>`);
    expect(r.doc.section.pageWidthPx).toBe(816);
  });

  it("yields one empty paragraph for an empty body (caret needs a home)", () => {
    const r = importBody(``);
    expect(r.doc.blocks).toHaveLength(1);
    expect(para(r.doc.blocks[0]).runs[0]!.text).toBe("");
  });

  it("locates a non-standard main part via [Content_Types].xml", () => {
    const bytes = makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_XML.replace("/word/document.xml", "/word/document2.xml"),
      "word/document2.xml": documentXml(`<w:p><w:r><w:t>moved</w:t></w:r></w:p>`),
    });
    const r = runImport(bytes);
    expect(para(r.doc.blocks[0]).runs[0]!.text).toBe("moved");
  });

  it("falls back to word/document.xml without content types", () => {
    const bytes = makeDocx({ "word/document.xml": documentXml(`<w:p><w:r><w:t>ok</w:t></w:r></w:p>`) });
    expect(para(runImport(bytes).doc.blocks[0]).runs[0]!.text).toBe("ok");
  });

  it("reports NO_DOCUMENT_PART for a zip without a document", () => {
    const bytes = makeDocx({ "hello.txt": "not a docx" });
    try {
      runImport(bytes);
      expect.unreachable();
    } catch (e) {
      expect(e).toBeInstanceOf(ImportError);
      expect((e as ImportError).code).toBe("NO_DOCUMENT_PART");
    }
  });

  it("mints ids that cannot collide with editor-minted ids", () => {
    const r = importBody(`<w:p><w:r><w:t>x</w:t></w:r></w:p>`);
    for (const b of r.doc.blocks) expect(b.id).toMatch(/^i\d+$/);
  });
});

describe("docx pipeline — real Word output", () => {
  // Generated by desktop Word (COM automation): bold/italic/size runs, center
  // alignment, a tab, and a 2x2 table. Word's default font (Aptos via the
  // minorHAnsi theme slot) and size live in styles.xml docDefaults — their
  // presence here proves the StyleResolver + theme wiring on real Word output.
  it("imports a Word-authored document", () => {
    const bytes = new Uint8Array(readFileSync(new URL("./__fixtures__/real-word-basic.docx", import.meta.url)));
    const r = runImport(bytes);

    const paras = r.doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
    expect(paras[0]!.runs[0]).toMatchObject({ text: "Bold heading", style: { bold: true } });
    expect(paras[1]!.runs[0]!.style).toMatchObject({ italic: true, fontSizePx: 18.67 }); // 14pt direct
    expect(paras[2]!).toMatchObject({ style: { align: "center" } });
    // docDefaults → theme minorHAnsi → Aptos; default size 11pt → 14.67px.
    expect(paras[2]!.runs[0]!.style.fontFamily).toBe("Aptos, serif");
    expect(paras[2]!.runs[0]!.style.fontSizePx).toBeCloseTo(14.67, 2);

    const tbl = r.doc.blocks.find((b): b is TableBlock => b.kind === "table");
    expect(tbl).toBeDefined();
    expect(tbl!.rows).toHaveLength(2);
    expect((tbl!.rows[0]!.cells[0]!.blocks[0] as Paragraph).runs[0]!.text).toBe("A1");

    expect(r.warnings.map((w) => w.code)).toEqual(["tabs"]);
  });
});
