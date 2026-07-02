// Ground-truth import test for issue #147: OOXML paragraph shading is
// tri-state — absent (inherit), present-with-fill (use it), and present-but-
// EMPTY (`<w:shd w:val="clear" w:fill="auto"/>` — Word's "Shading → No Color",
// an explicit override that clears shading inherited from a named style).
// A round-trip test cannot catch the third state on its own: the importer used
// to collapse it into "absent" and the resulting (wrong) model is a stable
// fixed point. The import assertions here are against the SOURCE XML's meaning;
// the export section then proves the clear is preserved through a write.

import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import type { Document, Paragraph } from "@cw/shared";
import { writeDocx } from "../../export/docx/writeDocx";
import { stylesPartXml, styledDocx } from "./fixture";
import { runImport } from "./pipeline";

// A paragraph style carrying shading (green fill), like a "Callout" style.
const STYLES = stylesPartXml(
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
     <w:name w:val="Normal"/>
   </w:style>
   <w:style w:type="paragraph" w:styleId="Callout">
     <w:basedOn w:val="Normal"/>
     <w:pPr><w:shd w:val="clear" w:color="auto" w:fill="D9EAD3"/></w:pPr>
   </w:style>`,
);

const importStyled = (body: string) => runImport(styledDocx(body, STYLES));

const para = (b: unknown): Paragraph => {
  expect((b as Paragraph).kind).toBe("paragraph");
  return b as Paragraph;
};

const exportedDocumentXml = (doc: Document): string =>
  strFromU8(unzipSync(writeDocx(doc).bytes)["word/document.xml"]!);

describe("w:shd tri-state (issue #147)", () => {
  it("inherits the style's shading when the paragraph has no w:shd (sanity)", () => {
    const r = importStyled(
      `<w:p><w:pPr><w:pStyle w:val="Callout"/></w:pPr><w:r><w:t>shaded</w:t></w:r></w:p>`,
    );
    expect(para(r.doc.blocks[0]).style.shading).toBe("#d9ead3");
  });

  it("uses the paragraph's own fill when it carries a direct w:shd (sanity)", () => {
    const r = importStyled(
      `<w:p><w:pPr><w:pStyle w:val="Callout"/><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/></w:pPr><w:r><w:t>override</w:t></w:r></w:p>`,
    );
    expect(para(r.doc.blocks[0]).style.shading).toBe("#fff2cc");
  });

  it('clears inherited shading on an explicit "No Color" (<w:shd w:val="clear" w:fill="auto"/>)', () => {
    // Word writes this exact element when the user picks Shading → No Color on
    // a paragraph whose STYLE carries a fill: present-but-empty = "no shading
    // here, overriding the style".
    const r = importStyled(
      `<w:p><w:pPr><w:pStyle w:val="Callout"/><w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:pPr><w:r><w:t>cleared</w:t></w:r></w:p>`,
    );
    const p = para(r.doc.blocks[0]);
    expect(p.style.shading).toBeUndefined();
    expect(p.style.shadingCleared).toBe(true); // the explicit-clear marker
  });

  it('clears inherited shading on a nil w:shd (<w:shd w:val="nil"/>)', () => {
    // The other spelling of "explicitly none" some producers emit.
    const r = importStyled(
      `<w:p><w:pPr><w:pStyle w:val="Callout"/><w:shd w:val="nil"/></w:pPr><w:r><w:t>cleared</w:t></w:r></w:p>`,
    );
    const p = para(r.doc.blocks[0]);
    expect(p.style.shading).toBeUndefined();
    expect(p.style.shadingCleared).toBe(true);
  });
});

describe("w:shd explicit-clear export (issue #147)", () => {
  it("re-emits an empty w:shd for a cleared paragraph under a shaded style", () => {
    const cleared = importStyled(
      `<w:p><w:pPr><w:pStyle w:val="Callout"/><w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:pPr><w:r><w:t>cleared</w:t></w:r></w:p>`,
    ).doc;
    const xml = exportedDocumentXml(cleared);
    // The paragraph's own w:shd is the empty "No Color" form, not the style fill.
    expect(xml).toContain(`<w:shd w:val="clear" w:color="auto" w:fill="auto"/>`);
    expect(xml).not.toContain("D9EAD3");
    expect(xml).not.toContain("d9ead3");
  });

  it("survives import → export → import (the clear is not lost)", () => {
    const a = importStyled(
      `<w:p><w:pPr><w:pStyle w:val="Callout"/><w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:pPr><w:r><w:t>cleared</w:t></w:r></w:p>`,
    ).doc;
    const b = runImport(writeDocx(a).bytes).doc;
    const p = para(b.blocks[0]);
    expect(p.style.shading).toBeUndefined();
    expect(p.style.shadingCleared).toBe(true);
  });

  it("a plain unshaded paragraph emits no w:shd (no false clear)", () => {
    const doc = importStyled(`<w:p><w:r><w:t>plain</w:t></w:r></w:p>`).doc;
    expect(exportedDocumentXml(doc)).not.toContain("<w:shd");
  });
});
