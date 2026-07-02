// Import + round-trip test for issue #160: Word's "automatic" paragraph spacing
// (w:beforeAutospacing / w:afterAutospacing — the HTML-<p> behavior, common in
// web-sourced docs). When on, Word ignores the explicit w:before/@w:after and
// computes the spacing itself; the importer used to drop the attributes, so an
// autospaced paragraph collapsed to its explicit value (often 0) on import. The
// model now bakes an approximate auto value and keeps a flag so export re-emits
// the attributes.

import { strFromU8, unzipSync } from "fflate";
import { AUTO_PARA_SPACING_PX, type Document, type Paragraph } from "@cw/shared";
import { describe, expect, it } from "vitest";
import { writeDocx } from "../../export/docx/writeDocx";
import { simpleDocx } from "./fixture";
import { runImport } from "./pipeline";

const para = (b: unknown): Paragraph => {
  expect((b as Paragraph).kind).toBe("paragraph");
  return b as Paragraph;
};
const exportedXml = (d: Document): string => strFromU8(unzipSync(writeDocx(d).bytes)["word/document.xml"]!);

// A paragraph with autospacing on and explicit before/after = 0 (what Word writes).
const AUTOSPACED = `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:beforeAutospacing="1" w:afterAutospacing="1"/></w:pPr><w:r><w:t>auto</w:t></w:r></w:p>`;

describe("paragraph auto-spacing (issue #160)", () => {
  it("bakes an auto value instead of collapsing to the explicit 0", () => {
    const d = runImport(simpleDocx(AUTOSPACED)).doc;
    const s = para(d.blocks[0]).style;
    expect(s.spaceBeforeAuto).toBe(true);
    expect(s.spaceAfterAuto).toBe(true);
    expect(s.spaceBeforePx).toBe(AUTO_PARA_SPACING_PX);
    expect(s.spaceAfterPx).toBe(AUTO_PARA_SPACING_PX);
  });

  it("leaves a normal paragraph's spacing untouched (no autospacing)", () => {
    const d = runImport(simpleDocx(`<w:p><w:pPr><w:spacing w:before="240" w:after="0"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`)).doc;
    const s = para(d.blocks[0]).style;
    expect(s.spaceBeforeAuto).toBeUndefined();
    expect(s.spaceBeforePx).toBe(16); // 240 twips = 16px, unchanged
    expect(s.spaceAfterPx).toBe(0);
  });

  it("re-emits the autospacing attributes on export", () => {
    const xml = exportedXml(runImport(simpleDocx(AUTOSPACED)).doc);
    expect(xml).toContain(`w:beforeAutospacing="1"`);
    expect(xml).toContain(`w:afterAutospacing="1"`);
  });

  it("survives import → export → import (autospacing not lost)", () => {
    const a = runImport(simpleDocx(AUTOSPACED)).doc;
    const b = runImport(writeDocx(a).bytes).doc;
    const s = para(b.blocks[0]).style;
    expect(s.spaceBeforeAuto).toBe(true);
    expect(s.spaceAfterAuto).toBe(true);
    expect(s.spaceBeforePx).toBe(AUTO_PARA_SPACING_PX);
  });
});
