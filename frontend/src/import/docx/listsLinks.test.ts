// Milestone 4 (high-impact): numbering/lists, resolved hyperlinks.

import { describe, expect, it } from "vitest";
import type { Paragraph } from "@cw/shared";
import { runImport } from "./pipeline";
import { CONTENT_TYPES_XML, documentXml, makeDocx, REL_TYPES, relsXml } from "./fixture";

const para = (b: unknown): Paragraph => {
  expect((b as Paragraph).kind).toBe("paragraph");
  return b as Paragraph;
};
const codes = (r: ReturnType<typeof runImport>): string[] => r.warnings.map((w) => w.code);

const NUMBERING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/>
      <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/>
      <w:rPr><w:rFonts w:ascii="Symbol"/></w:rPr>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="10"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="11"><w:abstractNumId w:val="1"/></w:num>
  <w:num w:numId="12"><w:abstractNumId w:val="0"/>
    <w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>
</w:numbering>`;

/** document.xml + numbering.xml wired through rels. */
function docxWithNumbering(body: string): Uint8Array {
  return makeDocx({
    "[Content_Types].xml": CONTENT_TYPES_XML,
    "word/document.xml": documentXml(body),
    "word/_rels/document.xml.rels": relsXml([
      { id: "rN", type: REL_TYPES.numbering, target: "numbering.xml" },
    ]),
    "word/numbering.xml": NUMBERING,
  });
}

const listItem = (numId: string, ilvl: number, text: string): string =>
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>` +
  `<w:ind w:left="720" w:hanging="360"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

describe("lists — numbering.xml", () => {
  it("sets list membership and a decimal definition", () => {
    const r = runImport(docxWithNumbering(listItem("10", 0, "first")));
    const p = para(r.doc.blocks[0]);
    expect(p.style.list).toEqual({ listId: "10", level: 0 });
    const def = r.doc.lists?.["10"];
    expect(def).toBeDefined();
    expect(def!.levels[0]).toMatchObject({ format: "decimal", text: "%1.", start: 1 });
  });

  it("only emits definitions for referenced lists", () => {
    const r = runImport(docxWithNumbering(listItem("10", 0, "x")));
    expect(Object.keys(r.doc.lists ?? {})).toEqual(["10"]); // 11, 12 unused
  });

  it("avoids double-counting list indent (engine adds level indent)", () => {
    // Paragraph resolves to 720tw (48px); level also carries 720tw. Engine adds
    // them, so the paragraph's own indent must drop to 0.
    const r = runImport(docxWithNumbering(listItem("10", 0, "x")));
    const p = para(r.doc.blocks[0]);
    expect(p.style.indentLeftPx).toBe(0);
    expect(p.style.indentFirstLinePx).toBe(0);
    expect(r.doc.lists!["10"]!.levels[0]!.indentLeftPx).toBe(48);
    expect(r.doc.lists!["10"]!.levels[0]!.hangingPx).toBe(24);
  });

  it("maps nested levels with their own formats", () => {
    const r = runImport(docxWithNumbering(listItem("10", 1, "nested")));
    expect(para(r.doc.blocks[0]).style.list).toEqual({ listId: "10", level: 1 });
    expect(r.doc.lists!["10"]!.levels[1]).toMatchObject({ format: "lowerLetter", text: "%2)" });
  });

  it("maps bullet lists with the Symbol glyph normalized to Unicode", () => {
    const r = runImport(docxWithNumbering(listItem("11", 0, "bullet")));
    const def = r.doc.lists!["11"]!;
    expect(def.levels[0]!.format).toBe("bullet");
    expect(def.levels[0]!.bulletChar).toBe("•"); // U+F0B7 → U+2022
  });

  it("applies startOverride from w:num/w:lvlOverride", () => {
    const r = runImport(docxWithNumbering(listItem("12", 0, "x")));
    expect(r.doc.lists!["12"]!.levels[0]!.start).toBe(5);
  });

  it("w:numId 0 removes inherited numbering", () => {
    const body = `<w:p><w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`;
    const r = runImport(docxWithNumbering(body));
    expect(para(r.doc.blocks[0]).style.list).toBeUndefined();
  });

  it("numbers a list item only once when it contains soft breaks", () => {
    // Word: a Shift+Enter inside a list item is the SAME item — one number, the
    // wrapped line carries no marker. We split on w:br, so only the first
    // segment may keep list membership.
    const body =
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="10"/></w:numPr></w:pPr>` +
      `<w:r><w:t>line one</w:t><w:br/><w:t>line two</w:t></w:r></w:p>` +
      listItem("10", 0, "next item");
    const r = runImport(docxWithNumbering(body));
    const ps = r.doc.blocks.map(para);
    expect(ps[0]!.style.list).toEqual({ listId: "10", level: 0 }); // numbered
    expect(ps[1]!.style.list).toBeUndefined(); // continuation — no marker
    expect(ps[2]!.style.list).toEqual({ listId: "10", level: 0 }); // next item numbered
    // Continuation aligns under the item text (level indent restored).
    expect(ps[1]!.style.indentLeftPx).toBe(r.doc.lists!["10"]!.levels[0]!.indentLeftPx);
  });

  it("does not number the empty paragraph left by a trailing soft break", () => {
    const body =
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="10"/></w:numPr></w:pPr>` +
      `<w:r><w:t>item</w:t><w:br/></w:r></w:p>`;
    const r = runImport(docxWithNumbering(body));
    const ps = r.doc.blocks.map(para);
    expect(ps[0]!.style.list).toEqual({ listId: "10", level: 0 });
    expect(ps[1]!.runs.map((x) => x.text).join("")).toBe("");
    expect(ps[1]!.style.list).toBeUndefined(); // empty trailing line not numbered
  });

  it("warns when a list reference has no definition", () => {
    const r = runImport(docxWithNumbering(listItem("99", 0, "x")));
    expect(para(r.doc.blocks[0]).style.list).toBeUndefined();
    expect(codes(r)).toContain("list-missing");
  });
});

describe("hyperlinks — resolved through rels", () => {
  function docxWithLink(target: string, external = true): Uint8Array {
    return makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_XML,
      "word/document.xml": documentXml(
        `<w:p><w:hyperlink r:id="rL"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>site</w:t></w:r></w:hyperlink></w:p>`,
      ),
      "word/_rels/document.xml.rels": relsXml([
        { id: "rL", type: REL_TYPES.hyperlink, target, external },
      ]),
    });
  }

  it("resolves an external hyperlink r:id to its URL", () => {
    const r = runImport(docxWithLink("https://example.com/page"));
    expect(para(r.doc.blocks[0]).runs[0]!.style.link).toBe("https://example.com/page");
    expect(codes(r)).not.toContain("links-unresolved");
  });
});
