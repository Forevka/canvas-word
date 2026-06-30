// w:caps / w:smallCaps must survive a full .docx round-trip: export emits the
// toggle inside w:rPr, and re-import re-tags the run's CharStyle.caps/smallCaps.
import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, SectionProps } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import { writeDocx } from "./writeDocx";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#202124" };
const PARA = { align: "left" as const, lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

const doc: Document = {
  section: SECTION,
  blocks: [{
    kind: "paragraph", id: "p1", revision: 0, style: { ...PARA },
    runs: [
      { text: "all caps", style: { ...CHAR, caps: true } },
      { text: " and ", style: { ...CHAR } },
      { text: "small caps", style: { ...CHAR, smallCaps: true } },
      { text: " and plain", style: { ...CHAR } },
    ],
  }],
};

describe("caps / smallCaps .docx round-trip", () => {
  it("re-imports w:caps and w:smallCaps on the right runs (and leaves plain runs plain)", () => {
    const bytes = writeDocx(doc).bytes;
    const back = runImport(bytes).doc;
    const p = back.blocks[0] as Paragraph;

    const caps = p.runs.find((r) => r.text === "all caps");
    const small = p.runs.find((r) => r.text === "small caps");
    const plain = p.runs.find((r) => r.text === " and plain");

    expect(caps?.style.caps).toBe(true);
    expect(caps?.style.smallCaps).not.toBe(true);
    expect(small?.style.smallCaps).toBe(true);
    expect(small?.style.caps).not.toBe(true);
    // No toggle bleed onto the surrounding plain text.
    expect(plain?.style.caps).not.toBe(true);
    expect(plain?.style.smallCaps).not.toBe(true);
  });
});
