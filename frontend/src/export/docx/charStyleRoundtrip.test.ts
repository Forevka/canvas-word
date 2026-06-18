// Character styles must survive a full .docx round-trip: export emits a
// w:type="character" style and a w:rStyle reference on the run; re-import
// rebuilds the NamedStyle (type "character") and re-tags the run's charStyleId.
import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, SectionProps, Stylesheet } from "@cw/shared";
import { styleById, styleType } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import { writeDocx } from "./writeDocx";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#202124" };
const PARA = { align: "left" as const, lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

const sheet: Stylesheet = {
  defaultStyleId: "Normal",
  styles: [
    { id: "Normal", name: "Normal", type: "paragraph", char: { ...CHAR }, para: { ...PARA } },
    { id: "Emphasis", name: "Emphasis", type: "character", char: { italic: true, color: "#ff0000" }, para: {} },
  ],
};

// "fancy" carries the character style: its concrete formatting is baked (italic
// + red) AND it references the style id.
const styledRun: CharStyle = { ...CHAR, italic: true, color: "#ff0000", charStyleId: "Emphasis" };

const doc: Document = {
  section: SECTION,
  stylesheet: sheet,
  blocks: [{
    kind: "paragraph", id: "p1", revision: 0, style: { ...PARA, namedStyle: "Normal" },
    runs: [
      { text: "plain ", style: { ...CHAR } },
      { text: "fancy", style: styledRun },
      { text: " text", style: { ...CHAR } },
    ],
  }],
};

describe("character style .docx round-trip", () => {
  it("re-imports the character style (typed) and the run's charStyleId reference", () => {
    const bytes = writeDocx(doc).bytes;
    const back = runImport(bytes).doc;

    const emphasis = styleById(back.stylesheet!, "Emphasis");
    expect(emphasis).toBeTruthy();
    expect(styleType(emphasis!)).toBe("character");
    expect(emphasis!.char.italic).toBe(true);
    expect(emphasis!.char.color).toBe("#ff0000");
    expect(Object.keys(emphasis!.para)).toHaveLength(0); // character styles carry no pPr

    const p = back.blocks[0] as Paragraph;
    const fancy = p.runs.find((r) => r.text === "fancy");
    expect(fancy?.style.charStyleId).toBe("Emphasis");
    // No "everything bold/italic" XOR regression: the plain runs stay plain.
    expect(p.runs.find((r) => r.text.includes("plain"))?.style.italic).toBe(false);
  });
});
