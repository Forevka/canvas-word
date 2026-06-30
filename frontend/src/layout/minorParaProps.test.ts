// Layout effects of the minor paragraph props (issue #62): w:widowControl drives
// pagination (orphan/widow split rules), and w:textAlignment shifts the line's
// shared baseline within a tall line box. The canvas stub (chars × ½ font-size)
// makes both deterministic. MUST import the stub first, before the engine loads.
import "./test-canvas-setup";

import { describe, it, expect } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, SectionProps } from "@cw/shared";
import { createLayoutEngine } from "./engine";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000000" };
const PARA: ParaStyle = { align: "left", lineHeight: 1, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
// content box: 816-192 = 624 wide, 1056-192 = 864 tall; 16px chars at ½ size = 8px,
// so lineHeight 1 → 16px lines → exactly 54 lines per page.
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

let n = 0;
const para = (text: string, style: Partial<ParaStyle> = {}): Paragraph => ({
  kind: "paragraph", id: `t${n++}`, revision: 0, runs: [{ text, style: { ...CHAR } }], style: { ...PARA, ...style },
});

const layout = (d: Document) => createLayoutEngine().layout(d);
const doc = (blocks: Paragraph[]): Document => ({ section: SECTION, blocks });
const pagesOf = (tree: ReturnType<typeof layout>, id: string): number[] => {
  const out: number[] = [];
  for (const pg of tree.pages) if (pg.blocks.some((b) => b.blockId === id)) out.push(pg.index);
  return out;
};
const firstLineAscent = (tree: ReturnType<typeof layout>, id: string): number => {
  for (const pg of tree.pages) for (const b of pg.blocks) if (b.blockId === id) return b.lines[0]!.ascent;
  throw new Error(`block ${id} not placed`);
};

describe("widowControl drives the orphan/widow split rule", () => {
  // Measure the page's single-line capacity (independent of the stub's exact line
  // height), then fill capacity-1 lines so EXACTLY one line of room remains — the
  // boundary the widow/orphan rule governs.
  const capacity = layout(doc(Array.from({ length: 400 }, (_, i) => para(`probe${i}`)))).pages[0]!.blocks.length;
  const fillers = (): Paragraph[] => Array.from({ length: capacity - 1 }, (_, i) => para(`f${i}`));
  // ~30 words wraps to several lines in the 624px content box, so only its first
  // line could fit in the single remaining row.
  const MULTI_LINE = Array.from({ length: 30 }, () => "lorem").join(" ");

  it("default (ON) moves the whole paragraph rather than strand a widow line", () => {
    const tgt = para(MULTI_LINE);
    const tree = layout(doc([...fillers(), tgt]));
    // Widow/orphan control keeps it whole on the next page — never a lone first line.
    expect(pagesOf(tree, tgt.id)).toEqual([1]);
  });

  it("widowControl:false lets the paragraph split across the page boundary", () => {
    const tgt = para(MULTI_LINE, { widowControl: false });
    const tree = layout(doc([...fillers(), tgt]));
    // With the rule off, the first line stays on page 1 and the rest flows over.
    expect(pagesOf(tree, tgt.id)).toEqual([0, 1]);
  });
});

describe("textAlignment shifts the line baseline within a tall line box", () => {
  // lineHeight 2 makes the line box twice the natural text height, so the baseline
  // has room to move; "top" rides high, "bottom" low, "center"/"baseline" in between.
  const ascentFor = (v?: ParaStyle["textAlignment"]): number => {
    const p = para("hello world", v ? { lineHeight: 2, textAlignment: v } : { lineHeight: 2 });
    return firstLineAscent(layout(doc([p])), p.id);
  };

  it("orders top < center < bottom", () => {
    const top = ascentFor("top");
    const center = ascentFor("center");
    const bottom = ascentFor("bottom");
    expect(top).toBeLessThan(center);
    expect(center).toBeLessThan(bottom);
  });

  it("treats baseline (and absent) as the centered default", () => {
    expect(ascentFor("baseline")).toBe(ascentFor(undefined));
    expect(ascentFor("baseline")).toBe(ascentFor("center"));
  });
});
