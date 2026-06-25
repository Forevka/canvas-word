// RTL/bidi paragraph and run direction round-trip tests.
// Mirrors the structure of docx.test.ts: build a doc, export, inspect XML, re-import.

import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import type { Document, Paragraph } from "@cw/shared";
import type { ParaStyle } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import { simpleDocx } from "../../import/docx/fixture";
import { writeDocx } from "./writeDocx";
import { paraCoreXml, partialPPrXml } from "./styleProps";

const BASE_PARA: ParaStyle = {
  align: "left",
  lineHeight: 1.5,
  spaceBeforePx: 0,
  spaceAfterPx: 0,
  indentFirstLinePx: 0,
  indentLeftPx: 0,
};

const roundTrip = (doc: Document): Document => runImport(writeDocx(doc).bytes).doc;
const exportedDocumentXml = (doc: Document): string =>
  strFromU8(unzipSync(writeDocx(doc).bytes)["word/document.xml"]!);
const paras = (d: Document): Paragraph[] =>
  d.blocks.filter((b): b is Paragraph => b.kind === "paragraph");

describe("bidi round-trip", () => {
  it("exports <w:bidi/> for direction:'rtl' and re-imports as direction:'rtl'", () => {
    // Build a document with one RTL paragraph via import from XML.
    const imported = runImport(
      simpleDocx(
        `<w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>` +
          `<w:r><w:t>مرحبا</w:t></w:r></w:p>`,
      ),
    ).doc;
    const p = paras(imported)[0]!;
    // Import should set direction to rtl.
    expect(p.style.direction).toBe("rtl");

    // Export should contain <w:bidi/> in the paragraph properties.
    const xml = exportedDocumentXml(imported);
    expect(xml).toContain("<w:bidi/>");

    // Re-import should preserve direction.
    const roundTripped = roundTrip(imported);
    expect(paras(roundTripped)[0]!.style.direction).toBe("rtl");
  });

  it("exports <w:rtl/> for run rtl:true and re-imports as rtl:true", () => {
    // Build a document with one RTL-flagged run via import from XML.
    const imported = runImport(
      simpleDocx(
        `<w:p><w:r><w:rPr><w:rtl/></w:rPr><w:t>hello</w:t></w:r></w:p>`,
      ),
    ).doc;
    const run = paras(imported)[0]!.runs[0]!;
    expect(run.style.rtl).toBe(true);

    // Export should contain <w:rtl/>.
    const xml = exportedDocumentXml(imported);
    expect(xml).toContain("<w:rtl/>");

    // Re-import should preserve rtl:true.
    const roundTripped = roundTrip(imported);
    expect(paras(roundTripped)[0]!.runs[0]!.style.rtl).toBe(true);
  });

  it("imports <w:bidi/> in pPr as direction:'rtl' (snippet test)", () => {
    const doc = runImport(
      simpleDocx(
        `<w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>` +
          `<w:r><w:t>x</w:t></w:r></w:p>`,
      ),
    ).doc;
    expect(paras(doc)[0]!.style.direction).toBe("rtl");
    // alignment is preserved alongside direction
    expect(paras(doc)[0]!.style.align).toBe("right");
  });

  it("does not emit <w:bidi/> for default ltr paragraphs", () => {
    const doc = runImport(
      simpleDocx(`<w:p><w:r><w:t>hello</w:t></w:r></w:p>`),
    ).doc;
    const xml = exportedDocumentXml(doc);
    expect(xml).not.toContain("<w:bidi/>");
  });

  it("does not emit <w:rtl/> for runs without rtl:true", () => {
    const doc = runImport(
      simpleDocx(`<w:p><w:r><w:t>hello</w:t></w:r></w:p>`),
    ).doc;
    const xml = exportedDocumentXml(doc);
    expect(xml).not.toContain("<w:rtl/>");
  });

  it("preserves an explicit w:bidi=\"0\" / w:rtl=\"0\" as a direction/rtl override (not absent)", () => {
    // An explicit OFF must clear inherited RTL, so it must NOT be dropped on import.
    const doc = runImport(
      simpleDocx(
        `<w:p><w:pPr><w:bidi w:val="0"/></w:pPr>` +
          `<w:r><w:rPr><w:rtl w:val="0"/></w:rPr><w:t>hello</w:t></w:r></w:p>`,
      ),
    ).doc;
    const p = paras(doc)[0]!;
    expect(p.style.direction).toBe("ltr"); // explicit off, not undefined
    expect(p.runs[0]!.style.rtl).toBe(false);

    // Export must emit the OFF state (w:val="0") so it survives a round-trip.
    const xml = exportedDocumentXml(doc);
    expect(xml).toContain('<w:bidi w:val="0"/>');
    expect(xml).toContain('<w:rtl w:val="0"/>');
    const rt = paras(roundTrip(doc))[0]!;
    expect(rt.style.direction).toBe("ltr");
    expect(rt.runs[0]!.style.rtl).toBe(false);
  });

  it("emits <w:bidi/> for an RTL named paragraph style (partialPPrXml / paraCoreXml)", () => {
    // Named-style export goes through styleProps, not documentXml.pPrXml.
    expect(partialPPrXml({ direction: "rtl" })).toContain("<w:bidi/>");
    expect(partialPPrXml({})).not.toContain("<w:bidi/>");
    expect(paraCoreXml({ ...BASE_PARA, direction: "rtl" })).toContain("<w:bidi/>");
    expect(paraCoreXml(BASE_PARA)).not.toContain("<w:bidi/>");
  });

  it("w:bidi ordering: <w:bidi/> appears before <w:spacing/> in exported pPr", () => {
    const doc = runImport(
      simpleDocx(
        `<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
      ),
    ).doc;
    const xml = exportedDocumentXml(doc);
    const bidiPos = xml.indexOf("<w:bidi/>");
    const spacingPos = xml.indexOf("<w:spacing ");
    expect(bidiPos).toBeGreaterThan(-1);
    expect(spacingPos).toBeGreaterThan(-1);
    expect(bidiPos).toBeLessThan(spacingPos);
  });
});
