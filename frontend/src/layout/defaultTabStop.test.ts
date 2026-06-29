// w:defaultTabStop (settings.xml) drives the tab interval a `\t` advances by when
// it runs past the last explicit tab stop. The canvas stub makes glyph width =
// chars × ½ font-size, so the geometry is exact: "a" (one 16px char) is 8px wide,
// and the first default stop strictly past 8px is the interval itself.
import "./test-canvas-setup";

import { describe, it, expect } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, SectionProps } from "@cw/shared";
import { createLayoutEngine } from "./engine";
import type { InlineFragment } from "./layoutTree";

const CHAR: CharStyle = {
  fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false,
  underline: false, strikethrough: false, color: "#000000",
};
const PARA: ParaStyle = { align: "left", lineHeight: 1, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

const tabbedDoc = (defaultTabStopPx?: number): Document => {
  const p: Paragraph = { kind: "paragraph", id: "p0", revision: 0, runs: [{ text: "a\tb", style: { ...CHAR } }], style: { ...PARA } };
  return { section: SECTION, blocks: [p], ...(defaultTabStopPx !== undefined ? { defaultTabStopPx } : {}) };
};

/** The x (block-relative) of the fragment whose text is "b" — i.e. where the tab landed. */
const tabbedX = (doc: Document): number => {
  const tree = createLayoutEngine().layout(doc);
  const frags = tree.pages[0]!.blocks[0]!.lines.flatMap((l) => l.fragments) as InlineFragment[];
  const b = frags.find((f) => f.text === "b");
  expect(b, "expected a fragment for the text after the tab").toBeTruthy();
  return b!.x;
};

describe("layout — w:defaultTabStop interval", () => {
  it("falls back to Word's 0.5in (48px) default when the doc sets none", () => {
    expect(tabbedX(tabbedDoc())).toBeCloseTo(48, 1);
  });

  it("honors a custom default tab interval (a \\t past the last stop advances by it)", () => {
    expect(tabbedX(tabbedDoc(96))).toBeCloseTo(96, 1);
    expect(tabbedX(tabbedDoc(72))).toBeCloseTo(72, 1);
  });

  it("a non-positive override is ignored (keeps the 48px fallback)", () => {
    expect(tabbedX(tabbedDoc(0))).toBeCloseTo(48, 1);
  });
});
