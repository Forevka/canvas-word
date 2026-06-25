// Bidi unit + engine-integration tests. The canvas stub (chars × ½ font-size)
// gives deterministic widths; bidi levels come from pretext's real Unicode data,
// so Hebrew/Arabic classify correctly without a font engine.
import "./test-canvas-setup";

import { describe, it, expect } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, Run, SectionProps, TableBlock, TableCell } from "@cw/shared";
import { baseLevelFor, effectiveAlign, hasRtlChars, levelsFor, visualOrder } from "./bidi";
import { createLayoutEngine } from "./engine";
import { caretRect, hitTest, selectionRects } from "./geometry";
import { defaultListDefinition, DEFAULT_BULLET_LIST_ID } from "@cw/shared";

describe("visualOrder (UAX#9 L2)", () => {
  it("leaves an all-LTR line in logical order", () => {
    expect(visualOrder([0, 0, 0])).toEqual([0, 1, 2]);
  });
  it("reverses a pure-RTL run", () => {
    expect(visualOrder([1, 1, 1])).toEqual([2, 1, 0]);
  });
  it("keeps an LTR island upright inside RTL (number in Arabic)", () => {
    // levels: R R L  → the trailing LTR item stays last visually but the two RTL
    // items flip: visual = [item2, item1, item0]? No — L2 reverses level≥1 runs.
    // [1,1,2]: level2 run {2} reversed (single), then level≥1 run {0,1,2} reversed.
    expect(visualOrder([1, 1, 2])).toEqual([2, 1, 0]);
  });
  it("flips two RTL words but keeps an embedded LTR word readable", () => {
    // "heb1 LTR heb2" as single-char items at levels [1,2,1] → [2,1,0]
    expect(visualOrder([1, 2, 1])).toEqual([2, 1, 0]);
  });
});

describe("levelsFor / base direction", () => {
  it("returns null for all-LTR text under an LTR base (fast path)", () => {
    expect(levelsFor("hello world", 0)).toBeNull();
  });
  it("marks Hebrew runs RTL under an LTR base", () => {
    const lv = levelsFor("ab אב", 0)!; // "ab א ב"
    expect(lv).not.toBeNull();
    expect(lv[0]!).toBe(0); // 'a' LTR
    expect(lv[3]! & 1).toBe(1); // aleph RTL
  });
  it("honours an explicit RTL base even when the first strong char is Latin", () => {
    // /pretext-patch base-level override: 'A' first, but base forced RTL.
    const lv = levelsFor("Aא", 1)!;
    expect(lv).not.toBeNull();
    expect(lv[1]! & 1).toBe(1); // aleph stays RTL
  });
  it("baseLevelFor maps direction to embedding level", () => {
    expect(baseLevelFor("rtl")).toBe(1);
    expect(baseLevelFor("ltr")).toBe(0);
    expect(baseLevelFor(undefined)).toBe(0);
  });
  it("hasRtlChars detects Hebrew/Arabic only", () => {
    expect(hasRtlChars("hello")).toBe(false);
    expect(hasRtlChars("שלום")).toBe(true); // שלום
    expect(hasRtlChars("مرحبا")).toBe(true); // مرحبا
  });
});

describe("effectiveAlign", () => {
  it("swaps left/right under RTL, leaves center/justify", () => {
    expect(effectiveAlign("left", "rtl")).toBe("right");
    expect(effectiveAlign("right", "rtl")).toBe("left");
    expect(effectiveAlign("center", "rtl")).toBe("center");
    expect(effectiveAlign("justify", "rtl")).toBe("justify");
  });
  it("is a no-op under LTR", () => {
    expect(effectiveAlign("left", "ltr")).toBe("left");
    expect(effectiveAlign("left", undefined)).toBe("left");
  });
});

// --- engine integration ---------------------------------------------------

const CHAR: CharStyle = {
  fontFamily: "Georgia, serif",
  fontSizePx: 16,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#000000",
};
const PARA: ParaStyle = {
  align: "left",
  lineHeight: 1,
  spaceBeforePx: 0,
  spaceAfterPx: 0,
  indentFirstLinePx: 0,
  indentLeftPx: 0,
};
const SECTION: SectionProps = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
};

let nextId = 0;
const para = (runs: Run[], style: Partial<ParaStyle> = {}): Paragraph => ({
  kind: "paragraph",
  id: `t${nextId++}`,
  revision: 0,
  runs,
  style: { ...PARA, ...style },
});
const doc = (p: Paragraph): Document => ({ section: SECTION, blocks: [p] });
const layout = (d: Document) => createLayoutEngine().layout(d);
const fragsOf = (d: Document) => layout(d).pages[0]!.blocks[0]!.lines[0]!.fragments;

describe("engine bidi reorder", () => {
  it("places two RTL runs in reversed visual order under an RTL paragraph", () => {
    // Two Hebrew runs (different styles → two fragments). Logical: שלום(0) עולם(1).
    const p = para(
      [
        { text: "שלום", style: { ...CHAR, bold: true } }, // שלום
        { text: "עולם", style: CHAR }, // עולם
      ],
      { direction: "rtl" },
    );
    const frags = fragsOf(doc(p));
    const first = frags.find((f) => f.startOffset === 0)!; // שלום
    const second = frags.find((f) => f.startOffset === 4)!; // עולם
    // RTL: the logically-first run sits to the RIGHT (greater x).
    expect(first.x).toBeGreaterThan(second.x);
    expect(first.level! & 1).toBe(1);
  });

  it("right-aligns an RTL paragraph by default (effectiveAlign)", () => {
    const p = para([{ text: "שלום", style: CHAR }], { direction: "rtl" });
    const frag = fragsOf(doc(p))[0]!;
    // 624px content box; 4 chars × 8px = 32px → starts near the right edge.
    expect(frag.x).toBeGreaterThan(500);
  });

  it("leaves an LTR paragraph untouched (no level tag, logical order)", () => {
    const p = para([{ text: "hello world", style: CHAR }]);
    const frag = fragsOf(doc(p))[0]!;
    expect(frag.x).toBeCloseTo(0, 0);
    expect(frag.level).toBeUndefined();
  });

  it("reorders an embedded Hebrew run inside an LTR paragraph", () => {
    // "abc" (LTR) then two Hebrew runs; the Hebrew pair flips, abc stays first.
    const p = para([
      { text: "abc ", style: CHAR },
      { text: "אבג", style: { ...CHAR, bold: true } }, // אבג
      { text: "דהו", style: CHAR }, // דהו
    ]);
    const frags = fragsOf(doc(p));
    const abc = frags.find((f) => f.text.startsWith("abc"))!;
    const heb1 = frags.find((f) => f.startOffset === 4)!; // אבג
    const heb2 = frags.find((f) => f.startOffset === 7)!; // דהו
    expect(abc.x).toBeLessThan(heb1.x);
    expect(abc.x).toBeLessThan(heb2.x);
    // Within the Hebrew sequence, the logically-first (heb1) is to the RIGHT.
    expect(heb1.x).toBeGreaterThan(heb2.x);
  });
});

describe("bidi geometry (caret / hit-test / selection)", () => {
  const rtlDoc = () => {
    const p = para([{ text: "שלוםעולם", style: CHAR }], { direction: "rtl" }); // 8 RTL chars
    return { doc: doc(p), id: p.id };
  };

  it("places the caret at offset 0 to the RIGHT of the caret at the end", () => {
    const { doc: d, id } = rtlDoc();
    const tree = layout(d);
    const xStart = caretRect(tree, { blockId: id, offset: 0 })!.x;
    const xEnd = caretRect(tree, { blockId: id, offset: 8 })!.x;
    expect(xStart).toBeGreaterThan(xEnd); // RTL: logical start is on the right
  });

  it("round-trips caret x → hitTest → same offset in an RTL run", () => {
    const { doc: d, id } = rtlDoc();
    const tree = layout(d);
    for (let off = 0; off <= 8; off++) {
      const c = caretRect(tree, { blockId: id, offset: off })!;
      const hit = hitTest(tree, c.pageIndex, c.x, c.y + 2)!;
      expect(hit.blockId).toBe(id);
      expect(hit.offset).toBe(off);
    }
  });

  it("produces a selection rect for a sub-range of an RTL run", () => {
    const { doc: d, id } = rtlDoc();
    const tree = layout(d);
    const rects = selectionRects(tree, { anchor: { blockId: id, offset: 2 }, focus: { blockId: id, offset: 5 } });
    expect(rects.length).toBeGreaterThanOrEqual(1);
    expect(rects[0]!.width).toBeGreaterThan(0);
  });

  // Regression: a MULTI-run RTL paragraph that wraps to several lines. Each line
  // holds several fragments stored in VISUAL (reversed) order, so the flattened
  // line index must derive each line's logical start/end from the min/max offset
  // across its fragments — not fragments[0]/[last]. With the old first/last logic
  // an offset at the logical start of a line (the visual RIGHT edge in RTL) was
  // attributed to the PREVIOUS line: the caret jumped up a line and whole-paragraph
  // selection broke.
  const multiRunRtlDoc = () => {
    // Alternating styles keep adjacent words as SEPARATE fragments; spaces give
    // wrap opportunities so the paragraph spans multiple lines.
    const runs: Run[] = [];
    for (let i = 0; i < 40; i++) {
      runs.push({ text: "מילה ", style: i % 2 === 0 ? CHAR : { ...CHAR, bold: true } });
    }
    const p = para(runs, { direction: "rtl" });
    return { doc: doc(p), id: p.id };
  };

  it("attributes each offset to its own visual line on a multi-run RTL paragraph", () => {
    const { doc: d, id } = multiRunRtlDoc();
    const tree = layout(d);
    const pb = tree.pages[0]!.blocks[0]!;
    expect(pb.lines.length).toBeGreaterThan(1); // must actually wrap

    for (const line of pb.lines) {
      // Logical span of this line, regardless of visual fragment order.
      const lineStart = Math.min(...line.fragments.map((f) => f.startOffset));
      const lineEnd = Math.max(...line.fragments.map((f) => f.endOffset));
      const expectedY = pb.y + line.y;
      // An offset safely INSIDE the line must render its caret on THIS line.
      const mid = Math.floor((lineStart + lineEnd) / 2);
      const c = caretRect(tree, { blockId: id, offset: mid })!;
      expect(c.y).toBeCloseTo(expectedY, 1);
      // And the line's own logical start belongs to this line, not the one above.
      const cs = caretRect(tree, { blockId: id, offset: lineStart })!;
      expect(cs.y).toBeCloseTo(expectedY, 1);
    }
  });

  it("selects every line of a single-run RTL paragraph that wraps", () => {
    // One Arabic run, no embedded LTR → one fragment PER LINE, but it wraps. The
    // narrow column in the showcase's bidi table is the real-world trigger.
    const word = "كلمة ";
    const p = para([{ text: word.repeat(60), style: CHAR }], { direction: "rtl" });
    const tree = layout(doc(p));
    const pb = tree.pages[0]!.blocks[0]!;
    expect(pb.lines.length).toBeGreaterThan(1);
    const end = Math.max(...pb.lines.flatMap((l) => l.fragments.map((f) => f.endOffset)));
    const rects = selectionRects(tree, {
      anchor: { blockId: p.id, offset: 0 },
      focus: { blockId: p.id, offset: end },
    });
    const ys = new Set(rects.map((r) => Math.round(r.y)));
    expect(ys.size).toBe(pb.lines.length); // a highlight row for EVERY line
    for (const r of rects) expect(r.width).toBeGreaterThan(0);
  });

  it("selects the whole multi-run RTL paragraph across every line", () => {
    const { doc: d, id } = multiRunRtlDoc();
    const tree = layout(d);
    const pb = tree.pages[0]!.blocks[0]!;
    const end = Math.max(...pb.lines.flatMap((l) => l.fragments.map((f) => f.endOffset)));
    const rects = selectionRects(tree, {
      anchor: { blockId: id, offset: 0 },
      focus: { blockId: id, offset: end },
    });
    // One row of highlight per visual line (each line is a contiguous logical range).
    const ys = new Set(rects.map((r) => Math.round(r.y)));
    expect(ys.size).toBe(pb.lines.length);
  });
});

describe("RTL text inside a narrow table cell (wraps to multiple lines)", () => {
  // Mirrors the showcase's bidi table: a right-to-left Arabic cell in a 3-column
  // table, so the column is narrow enough to wrap the text onto several lines.
  const tableDoc = () => {
    const cellPara: Paragraph = {
      kind: "paragraph",
      id: `cell${nextId++}`,
      revision: 0,
      runs: [{ text: "كلمة ".repeat(12), style: CHAR }],
      style: { ...PARA, direction: "rtl" },
    };
    const mk = (txt: string): TableCell => ({
      id: `c${nextId++}`,
      blocks: [{ kind: "paragraph", id: `p${nextId++}`, revision: 0, runs: [{ text: txt, style: CHAR }], style: PARA }],
    });
    const table: TableBlock = {
      kind: "table",
      id: `tbl${nextId++}`,
      revision: 0,
      rows: [{ cells: [{ id: `c${nextId++}`, blocks: [cellPara] }, mk("ltr"), mk("ltr")] }],
    };
    return { doc: { section: SECTION, blocks: [table] } as Document, id: cellPara.id };
  };

  const cellBlock = (tree: ReturnType<typeof layout>, id: string) => {
    for (const page of tree.pages) {
      for (const block of page.blocks) {
        for (const row of block.table?.rows ?? []) {
          for (const c of row.cells) {
            const pb = c.blocks.find((b) => b.blockId === id);
            if (pb) return pb;
          }
        }
      }
    }
    throw new Error("cell block not found");
  };

  it("wraps the RTL cell text across more than one line", () => {
    const { doc: d, id } = tableDoc();
    const pb = cellBlock(layout(d), id);
    expect(pb.lines.length).toBeGreaterThan(1);
  });

  it("highlights EVERY wrapped line when the whole cell paragraph is selected", () => {
    const { doc: d, id } = tableDoc();
    const tree = layout(d);
    const pb = cellBlock(tree, id);
    const end = Math.max(...pb.lines.flatMap((l) => l.fragments.map((f) => f.endOffset)));
    const rects = selectionRects(tree, {
      anchor: { blockId: id, offset: 0 },
      focus: { blockId: id, offset: end },
    });
    const ys = new Set(rects.map((r) => Math.round(r.y)));
    expect(ys.size).toBe(pb.lines.length);
  });

  it("round-trips a click on each wrapped line back to that line", () => {
    const { doc: d, id } = tableDoc();
    const tree = layout(d);
    const pb = cellBlock(tree, id);
    for (const line of pb.lines) {
      const mid = Math.floor(
        (Math.min(...line.fragments.map((f) => f.startOffset)) +
          Math.max(...line.fragments.map((f) => f.endOffset))) / 2,
      );
      const c = caretRect(tree, { blockId: id, offset: mid })!;
      const hit = hitTest(tree, c.pageIndex, c.x, c.y + 2)!;
      expect(hit.blockId).toBe(id);
    }
  });
});

describe("RTL list marker + justification", () => {
  const listDoc = (direction?: "rtl") => {
    const p = para([{ text: "פריט ברשימה", style: CHAR }], {
      ...(direction ? { direction } : {}),
      list: { listId: DEFAULT_BULLET_LIST_ID, level: 0 },
    });
    return { doc: { ...doc(p), lists: { [DEFAULT_BULLET_LIST_ID]: defaultListDefinition("bullet") } }, id: p.id };
  };

  it("hangs the list marker to the RIGHT of the text in an RTL paragraph", () => {
    const tree = layout(listDoc("rtl").doc);
    const pb = tree.pages[0]!.blocks[0]!;
    const line = pb.lines[0]!;
    const textRight = pb.x + Math.max(...line.fragments.map((f) => f.x + f.width));
    expect(pb.marker).toBeDefined();
    // RTL: marker sits at/right of the text's right edge (not in the left gutter).
    expect(pb.marker!.x).toBeGreaterThan(textRight - 1);
  });

  it("keeps the LTR list marker to the LEFT of the text", () => {
    const tree = layout(listDoc().doc);
    const pb = tree.pages[0]!.blocks[0]!;
    expect(pb.marker!.x).toBeLessThan(pb.x);
  });

  it("justifies the non-last lines of a multi-line RTL paragraph to full width", () => {
    // Long Hebrew text forces several lines; justify should fill the content box.
    const word = "מילה";
    const p = para([{ text: Array.from({ length: 60 }, () => word).join(" "), style: CHAR }], {
      direction: "rtl",
      align: "left", // → start/end-aware: "left" under RTL maps to right, but justify wins
    });
    const pj = { ...p, style: { ...p.style, align: "justify" as const } };
    const tree = layout(doc(pj));
    const lines = tree.pages[0]!.blocks[0]!.lines;
    expect(lines.length).toBeGreaterThan(1);
    const contentWidth = 624; // 816 - 2*96
    // A justified non-last line fills (nearly) the full content width.
    const firstLineRight = Math.max(...lines[0]!.fragments.map((f) => f.x + f.width));
    expect(firstLineRight).toBeGreaterThan(contentWidth - 24);
  });
});
