// Ground-truth import + round-trip test for issue #153: paragraph borders are
// tri-state — absent (inherit the style's border box), present-with-edges (use
// them), and present-but-EMPTY (an empty `<w:pBdr/>` or all-nil edges — an
// explicit override that CLEARS a box inherited from a named style). The
// importer used to collapse the empty form into "absent", so a paragraph that
// cleared its style's box re-inherited it. Parity with the cell/table
// `bordersSpecified` handling. Asserted against the SOURCE XML's meaning plus an
// export re-emission check.

import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import type { Document, Paragraph } from "@cw/shared";
import { writeDocx } from "../../export/docx/writeDocx";
import { stylesPartXml, styledDocx } from "./fixture";
import { runImport } from "./pipeline";

// A paragraph style carrying a full border box, like a "Boxed" callout style.
const STYLES = stylesPartXml(
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
   <w:style w:type="paragraph" w:styleId="Boxed">
     <w:name w:val="Boxed"/><w:basedOn w:val="Normal"/>
     <w:pPr><w:pBdr>
       <w:top w:val="single" w:sz="12" w:color="1A73E8"/>
       <w:left w:val="single" w:sz="12" w:color="1A73E8"/>
       <w:bottom w:val="single" w:sz="12" w:color="1A73E8"/>
       <w:right w:val="single" w:sz="12" w:color="1A73E8"/>
     </w:pBdr></w:pPr>
   </w:style>`,
);
const importStyled = (body: string): Document => runImport(styledDocx(body, STYLES)).doc;
const para = (b: unknown): Paragraph => {
  expect((b as Paragraph).kind).toBe("paragraph");
  return b as Paragraph;
};
const exportedXml = (d: Document): string => strFromU8(unzipSync(writeDocx(d).bytes)["word/document.xml"]!);

describe("w:pBdr border-clear tri-state (issue #153)", () => {
  it("inherits the style's border box when the paragraph has no w:pBdr (sanity)", () => {
    const d = importStyled(`<w:p><w:pPr><w:pStyle w:val="Boxed"/></w:pPr><w:r><w:t>boxed</w:t></w:r></w:p>`);
    expect(para(d.blocks[0]).style.borders).toBeDefined();
    expect(para(d.blocks[0]).style.bordersCleared).toBeUndefined();
  });

  it("clears the inherited box on an explicit empty w:pBdr", () => {
    const d = importStyled(
      `<w:p><w:pPr><w:pStyle w:val="Boxed"/><w:pBdr/></w:pPr><w:r><w:t>no box</w:t></w:r></w:p>`,
    );
    const p = para(d.blocks[0]);
    expect(p.style.borders).toBeUndefined();
    expect(p.style.bordersCleared).toBe(true);
  });

  it("clears the inherited box on all-nil w:pBdr edges", () => {
    const d = importStyled(
      `<w:p><w:pPr><w:pStyle w:val="Boxed"/><w:pBdr><w:top w:val="nil"/></w:pBdr></w:pPr><w:r><w:t>no box</w:t></w:r></w:p>`,
    );
    const p = para(d.blocks[0]);
    expect(p.style.borders).toBeUndefined();
    expect(p.style.bordersCleared).toBe(true);
  });
});

describe("w:pBdr border-clear export/round-trip (issue #153)", () => {
  it("re-emits an empty w:pBdr for a cleared paragraph under a boxed style", () => {
    const d = importStyled(
      `<w:p><w:pPr><w:pStyle w:val="Boxed"/><w:pBdr/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
    );
    expect(exportedXml(d)).toContain("<w:pBdr/>");
    expect(exportedXml(d)).not.toContain("1A73E8"); // the style box did NOT bake back in
    expect(exportedXml(d)).not.toContain("1a73e8");
  });

  it("survives import → export → import (the clear is not lost)", () => {
    const a = importStyled(
      `<w:p><w:pPr><w:pStyle w:val="Boxed"/><w:pBdr/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
    );
    const b = runImport(writeDocx(a).bytes).doc;
    const p = para(b.blocks[0]);
    expect(p.style.borders).toBeUndefined();
    expect(p.style.bordersCleared).toBe(true);
  });
});
