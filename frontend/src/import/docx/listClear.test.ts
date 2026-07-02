// Ground-truth import + round-trip test for issue #152: list membership is
// tri-state — absent (inherit the style's list), present-with-numId (use it),
// and explicit REMOVAL (`<w:numPr><w:numId w:val="0"/></w:numPr>` — Word's
// numId=0 sentinel, written when you toggle the list off on a list-styled
// paragraph). The importer used to collapse the removal into "absent", so a
// de-listed paragraph re-inherited the style's list. Asserted against the
// SOURCE XML's meaning plus an export re-emission check.

import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import type { Document, Paragraph } from "@cw/shared";
import { writeDocx } from "../../export/docx/writeDocx";
import { CONTENT_TYPES_XML, documentXml, makeDocx, REL_TYPES, relsXml } from "./fixture";
import { runImport } from "./pipeline";

// A paragraph style that carries numbering (numId 1), plus a real numbering part.
const STYLES = `<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="ListStyle">
    <w:name w:val="List Style"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
  </w:style>
</w:styles>`;
const NUMBERING = `<?xml version="1.0"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;

const importDoc = (body: string): Document =>
  runImport(
    makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_XML,
      "word/document.xml": documentXml(body),
      "word/_rels/document.xml.rels": relsXml([
        { id: "rS", type: REL_TYPES.styles, target: "styles.xml" },
        { id: "rN", type: REL_TYPES.numbering, target: "numbering.xml" },
      ]),
      "word/styles.xml": STYLES,
      "word/numbering.xml": NUMBERING,
    }),
  ).doc;

const para = (b: unknown): Paragraph => {
  expect((b as Paragraph).kind).toBe("paragraph");
  return b as Paragraph;
};
const exportedXml = (d: Document): string => strFromU8(unzipSync(writeDocx(d).bytes)["word/document.xml"]!);

describe("w:numPr list-removal tri-state (issue #152)", () => {
  it("inherits the style's list when the paragraph has no w:numPr (sanity)", () => {
    const d = importDoc(`<w:p><w:pPr><w:pStyle w:val="ListStyle"/></w:pPr><w:r><w:t>item</w:t></w:r></w:p>`);
    expect(para(d.blocks[0]).style.list).toBeDefined();
    expect(para(d.blocks[0]).style.listCleared).toBeUndefined();
  });

  it("removes the inherited list on an explicit numId=0", () => {
    const d = importDoc(
      `<w:p><w:pPr><w:pStyle w:val="ListStyle"/><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>not a list item</w:t></w:r></w:p>`,
    );
    const p = para(d.blocks[0]);
    expect(p.style.list).toBeUndefined();
    expect(p.style.listCleared).toBe(true);
  });
});

describe("w:numPr list-removal export/round-trip (issue #152)", () => {
  it("re-emits w:numId=0 for a de-listed paragraph under a list style", () => {
    const d = importDoc(
      `<w:p><w:pPr><w:pStyle w:val="ListStyle"/><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
    );
    expect(exportedXml(d)).toContain(`<w:numId w:val="0"/>`);
  });

  it("survives import → export → import (the opt-out is not lost)", () => {
    const a = importDoc(
      `<w:p><w:pPr><w:pStyle w:val="ListStyle"/><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
    );
    const b = runImport(writeDocx(a).bytes).doc;
    const p = para(b.blocks[0]);
    expect(p.style.list).toBeUndefined(); // did NOT re-inherit the style's list
    expect(p.style.listCleared).toBe(true);
  });
});
