// Ground-truth import + round-trip test for issue #154: a paragraph can REMOVE a
// specific tab stop inherited from its style with `<w:tab w:val="clear" w:pos="…"/>`.
// The importer used to discard "clear" tabs entirely at parse, so the inherited
// stop survived and re-emitted — the removal was silently lost. Asserted against
// the SOURCE XML's meaning plus an export re-emission check.

import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import type { Document, Paragraph } from "@cw/shared";
import { writeDocx } from "../../export/docx/writeDocx";
import { stylesPartXml, styledDocx } from "./fixture";
import { runImport } from "./pipeline";

// A paragraph style that defines a tab stop at 1440 twips (1"), so a paragraph
// can inherit it — or explicitly clear it.
const STYLES = stylesPartXml(
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
   <w:style w:type="paragraph" w:styleId="Tabbed">
     <w:name w:val="Tabbed"/><w:basedOn w:val="Normal"/>
     <w:pPr><w:tabs><w:tab w:val="left" w:pos="1440"/></w:tabs></w:pPr>
   </w:style>`,
);
const importStyled = (body: string): Document => runImport(styledDocx(body, STYLES)).doc;
const para = (b: unknown): Paragraph => {
  expect((b as Paragraph).kind).toBe("paragraph");
  return b as Paragraph;
};
const exportedXml = (d: Document): string => strFromU8(unzipSync(writeDocx(d).bytes)["word/document.xml"]!);

describe("w:tab clear tri-state (issue #154)", () => {
  it("inherits the style's tab stop when the paragraph has no w:tabs (sanity)", () => {
    const d = importStyled(`<w:p><w:pPr><w:pStyle w:val="Tabbed"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`);
    const stops = para(d.blocks[0]).style.tabStops ?? [];
    expect(stops.some((t) => !t.cleared && Math.round(t.posPx) === 96)).toBe(true); // 1440 twips = 96px
  });

  it("records a cleared stop for an explicit w:tab w:val=clear (not dropped)", () => {
    const d = importStyled(
      `<w:p><w:pPr><w:pStyle w:val="Tabbed"/><w:tabs><w:tab w:val="clear" w:pos="1440"/></w:tabs></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
    );
    const stops = para(d.blocks[0]).style.tabStops ?? [];
    const cleared = stops.find((t) => t.cleared);
    expect(cleared).toBeDefined();
    expect(Math.round(cleared!.posPx)).toBe(96);
    // No live (non-cleared) stop survives at that position — the tab is removed.
    expect(stops.some((t) => !t.cleared && Math.round(t.posPx) === 96)).toBe(false);
  });
});

describe("w:tab clear export/round-trip (issue #154)", () => {
  it("re-emits w:val=clear for a cleared tab under a tabbed style", () => {
    const d = importStyled(
      `<w:p><w:pPr><w:pStyle w:val="Tabbed"/><w:tabs><w:tab w:val="clear" w:pos="1440"/></w:tabs></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
    );
    expect(exportedXml(d)).toContain(`<w:tab w:val="clear" w:pos="1440"/>`);
  });

  it("survives import → export → import (the clear is not lost)", () => {
    const a = importStyled(
      `<w:p><w:pPr><w:pStyle w:val="Tabbed"/><w:tabs><w:tab w:val="clear" w:pos="1440"/></w:tabs></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
    );
    const b = runImport(writeDocx(a).bytes).doc;
    const stops = para(b.blocks[0]).style.tabStops ?? [];
    expect(stops.some((t) => t.cleared && Math.round(t.posPx) === 96)).toBe(true);
    expect(stops.some((t) => !t.cleared && Math.round(t.posPx) === 96)).toBe(false);
  });
});
