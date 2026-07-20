// OOXML color values must be 6 hex digits or the literal "auto"; Word rejects a
// 3-digit shorthand like "fff". The showcase uses #fff (white text), so exported
// files carried w:color w:val="fff" and Word refused to open them (issue #193).
import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import type { CharStyle, Document, Paragraph, ParaStyle, SectionProps } from "@cw/shared";
import { hexColor } from "./mappings";
import { writeDocx } from "./writeDocx";

describe("hexColor normalization (issue #193)", () => {
  it("expands 3-digit shorthand to 6 digits", () => {
    expect(hexColor("#fff")).toBe("ffffff");
    expect(hexColor("#0af")).toBe("00aaff");
    expect(hexColor("abc")).toBe("aabbcc");
  });
  it("passes 6-digit and auto through unchanged (lowercased)", () => {
    expect(hexColor("#1A2B3C")).toBe("1a2b3c");
    expect(hexColor("auto")).toBe("auto");
  });
});

describe("exported colors are schema-valid 6-hex values", () => {
  const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 14, bold: false, italic: false, underline: false, strikethrough: false, color: "#fff" };
  const PARA: ParaStyle = { align: "left", lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
  const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

  it("a #fff run color exports as w:val=\"ffffff\", never \"fff\"", () => {
    const doc: Document = {
      section: SECTION,
      blocks: [{ kind: "paragraph", id: "p1", revision: 0, runs: [{ text: "white", style: { ...CHAR } }], style: { ...PARA, shading: "#fff" } } as Paragraph],
    };
    const xml = strFromU8(unzipSync(writeDocx(doc).bytes)["word/document.xml"]!);
    expect(xml).toContain('<w:color w:val="ffffff"/>');
    expect(xml).not.toContain('w:val="fff"');
    expect(xml).not.toContain('w:fill="fff"');
  });
});
