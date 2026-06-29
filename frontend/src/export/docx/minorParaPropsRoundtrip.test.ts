// Minor paragraph properties (issue #62): w:widowControl, w:suppressLineNumbers,
// w:textAlignment, w:mirrorIndents, w:adjustRightInd must survive a full .docx
// round-trip and be parsed from a hand-written w:pPr. The importer is the oracle:
// a model → writeDocx → re-import that drops any of them would surface here.
import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import type { CharStyle, Document, Paragraph, ParaStyle, SectionProps } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import { simpleDocx } from "../../import/docx/fixture";
import { writeDocx } from "./writeDocx";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 14, bold: false, italic: false, underline: false, strikethrough: false, color: "#202124" };
const PARA: ParaStyle = { align: "left", lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

let n = 0;
const para = (text: string, style: Partial<ParaStyle> = {}): Paragraph => ({
  kind: "paragraph", id: `p${n++}`, revision: 0, runs: [{ text, style: { ...CHAR } }], style: { ...PARA, ...style },
});

const roundTrip = (doc: Document): Document => runImport(writeDocx(doc).bytes).doc;
const exportedXml = (doc: Document): string => strFromU8(unzipSync(writeDocx(doc).bytes)["word/document.xml"]!);
const paras = (d: Document): Paragraph[] => d.blocks.filter((b): b is Paragraph => b.kind === "paragraph");

describe("minor paragraph properties — w:widowControl / w:suppressLineNumbers / w:textAlignment / w:mirrorIndents / w:adjustRightInd", () => {
  it("emits each element with the right on/off encoding", () => {
    const xml = exportedXml({
      section: SECTION,
      blocks: [
        para("off", { widowControl: false, suppressLineNumbers: true, mirrorIndents: true, adjustRightInd: true, textAlignment: "bottom" }),
        para("on", { widowControl: true }),
        para("plain"),
      ],
    });
    // widowControl OFF carries an explicit w:val="0"; ON is a bare element.
    expect(xml).toContain('<w:widowControl w:val="0"/>');
    expect(xml).toContain("<w:widowControl/>");
    expect(xml).toContain("<w:suppressLineNumbers/>");
    expect(xml).toContain("<w:mirrorIndents/>");
    expect(xml).toContain("<w:adjustRightInd/>");
    expect(xml).toContain('<w:textAlignment w:val="bottom"/>');
  });

  it("preserves every minor prop through a full export → re-import round-trip", () => {
    const out = roundTrip({
      section: SECTION,
      blocks: [
        para("a", { widowControl: false, suppressLineNumbers: true, textAlignment: "top", mirrorIndents: true, adjustRightInd: true }),
        para("b", { widowControl: true, textAlignment: "center" }),
        para("c"),
      ],
    });
    const [a, b, c] = paras(out);
    expect(a!.style.widowControl).toBe(false);
    expect(a!.style.suppressLineNumbers).toBe(true);
    expect(a!.style.textAlignment).toBe("top");
    expect(a!.style.mirrorIndents).toBe(true);
    expect(a!.style.adjustRightInd).toBe(true);
    expect(b!.style.widowControl).toBe(true);
    expect(b!.style.textAlignment).toBe("center");
    // A plain paragraph keeps the defaults absent (no spurious round-trip values).
    expect(c!.style.widowControl).toBeUndefined();
    expect(c!.style.suppressLineNumbers).toBeUndefined();
    expect(c!.style.textAlignment).toBeUndefined();
    expect(c!.style.mirrorIndents).toBeUndefined();
    expect(c!.style.adjustRightInd).toBeUndefined();
  });

  it("parses the props from a hand-written w:pPr", () => {
    const body =
      `<w:p><w:pPr>` +
      `<w:widowControl w:val="0"/><w:suppressLineNumbers/><w:mirrorIndents/><w:adjustRightInd/>` +
      `<w:textAlignment w:val="center"/>` +
      `</w:pPr><w:r><w:t>hand</w:t></w:r></w:p>`;
    const p = paras(runImport(simpleDocx(body)).doc)[0]!;
    expect(p.style.widowControl).toBe(false);
    expect(p.style.suppressLineNumbers).toBe(true);
    expect(p.style.mirrorIndents).toBe(true);
    expect(p.style.adjustRightInd).toBe(true);
    expect(p.style.textAlignment).toBe("center");
  });

  it("ignores w:textAlignment='auto' (Word's default) instead of inventing a value", () => {
    const body = `<w:p><w:pPr><w:textAlignment w:val="auto"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`;
    const p = paras(runImport(simpleDocx(body)).doc)[0]!;
    expect(p.style.textAlignment).toBeUndefined();
  });
});
