import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, Run, SectionProps } from "@cw/shared";
import { linkAtPosition } from "./commands";

// linkAtPosition — the position-based hyperlink lookup that drives the hyperlink
// context toolbar. Links live on run styles (run.style.link).

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 14, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA: ParaStyle = { align: "left", lineHeight: 1.35, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

const run = (text: string, link?: string): Run => ({ text, style: { ...CHAR, ...(link ? { link } : {}) } });
// "see " [0,4)  ·  "here" [4,8) linked  ·  " now" [8,12)
const para: Paragraph = { kind: "paragraph", id: "p1", revision: 0, runs: [run("see "), run("here", "https://x.dev"), run(" now")], style: { ...PARA } };
const doc: Document = { section: SECTION, blocks: [para] };
const at = (offset: number): string | null => linkAtPosition(doc, { blockId: "p1", offset });

describe("linkAtPosition", () => {
  it("returns the URL when the caret is inside the linked run", () => {
    expect(at(5)).toBe("https://x.dev");
    expect(at(6)).toBe("https://x.dev");
  });

  it("returns null when the caret is in an unlinked run", () => {
    expect(at(0)).toBeNull();
    expect(at(2)).toBeNull();
    expect(at(10)).toBeNull();
  });

  it("treats the leading edge of the link as linked (entering it)", () => {
    expect(at(4)).toBe("https://x.dev");
  });

  it("treats the trailing edge of the link as linked (still touching it)", () => {
    expect(at(8)).toBe("https://x.dev");
  });

  it("returns null at the very end of the paragraph and for an unknown block", () => {
    expect(at(12)).toBeNull();
    expect(linkAtPosition(doc, { blockId: "nope", offset: 0 })).toBeNull();
  });
});
