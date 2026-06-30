// Regression: a single ArrowLeft/Right press must always visibly move the caret,
// even across a {page}/{pages} token field (whose model range collapses onto one
// glyph) or a hidden (w:vanish) run (which lays out zero-width). Before the fix
// the caret stepped one MODEL offset per press and appeared frozen for several
// presses while crossing those runs.
//
// `advanceVisibleStop` is the pure per-paragraph stepper the controller's caret
// movement is built on; we drive it against a real layout tree (via caretRect) so
// the collapse behaviour under test is the genuine engine geometry.
import "../layout/test-canvas-setup";

import { describe, it, expect } from "vitest";
import type { CharStyle, Document, FieldDef, Paragraph, ParaStyle, Run, SectionProps } from "@cw/shared";
import { createLayoutEngine } from "../layout/engine";
import { caretRect } from "../layout/geometry";
import { advanceVisibleStop, type CaretXY } from "./selectionController";

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

const para = (runs: Run[]): Paragraph => ({ kind: "paragraph", id: "p0", revision: 0, runs, style: { ...PARA } });
const text = (runs: Run[]): string => runs.map((r) => r.text).join("");

/** Build a one-paragraph doc and a (text, caretXY) probe over its real layout. */
function probe(runs: Run[], extra: Partial<Document> = {}) {
  const doc: Document = { section: SECTION, blocks: [para(runs)], ...extra };
  const tree = createLayoutEngine().layout(doc);
  const caretXY = (offset: number): CaretXY | null => {
    const r = caretRect(tree, { blockId: "p0", offset });
    return r ? { page: r.pageIndex, x: r.x, y: r.y } : null;
  };
  return { txt: text(runs), caretXY, x: (offset: number) => caretXY(offset)!.x };
}

const pageRuns = (): Run[] => {
  const fid = "f1";
  return [
    { text: "ab", style: CHAR },
    { text: "{page}", style: { ...CHAR, fieldId: fid } }, // model offsets 2..8
    { text: "cd", style: CHAR },
  ];
};
const pageFields: Record<string, FieldDef> = {
  f1: { id: "f1", instruction: " PAGE ", name: "PAGE", kind: "builtin", spec: { type: "PAGE" } },
};

describe("advanceVisibleStop — skips visually-collapsed runs", () => {
  it("ArrowRight crosses a {page} field with no dead presses", () => {
    const { txt, caretXY, x } = probe(pageRuns(), { fields: pageFields });
    let off = 2; // caret immediately before the field
    let prevX = x(off);
    for (let i = 0; i < 3; i++) {
      const next = advanceVisibleStop(txt, off, 1, false, caretXY);
      expect(next).not.toBeNull();
      off = next!;
      expect(x(off)).toBeGreaterThan(prevX); // every press visibly advances the caret
      prevX = x(off);
    }
    expect(off).toBeGreaterThanOrEqual(8); // walked past the field's model range
  });

  it("ArrowLeft crosses a {page} field with no dead presses", () => {
    const { txt, caretXY, x } = probe(pageRuns(), { fields: pageFields });
    let off = 10; // caret at the end ("...cd")
    let prevX = x(off);
    for (let i = 0; i < 3; i++) {
      const next = advanceVisibleStop(txt, off, -1, false, caretXY);
      expect(next).not.toBeNull();
      off = next!;
      expect(x(off)).toBeLessThan(prevX);
      prevX = x(off);
    }
    expect(off).toBeLessThanOrEqual(2); // reached the field's start or before
  });

  it("one ArrowLeft jumps over a hidden run preceding a bookmarked run", () => {
    // Mirrors sampleDoc: "abc" + hidden "HIDDEN" + "xyz" (the bookmarked text).
    const runs: Run[] = [
      { text: "abc", style: CHAR },
      { text: "HIDDEN", style: { ...CHAR, hidden: true } }, // model offsets 3..9, zero-width
      { text: "xyz", style: CHAR },
    ];
    const { txt, caretXY, x } = probe(runs, {
      bookmarks: { sample: { start: { blockId: "p0", offset: 9 }, end: { blockId: "p0", offset: 12 } } },
    });
    const start = 9; // immediately before the bookmark (right after the hidden run)
    const next = advanceVisibleStop(txt, start, -1, false, caretXY);
    expect(next).not.toBeNull();
    expect(x(next!)).toBeLessThan(x(start)); // a single press moves
    expect(next!).toBeLessThanOrEqual(3); // skipped the whole hidden run at once
  });

  it("runs off the block edge (returns null) when the whole suffix is collapsed", () => {
    const runs: Run[] = [
      { text: "ab", style: CHAR },
      { text: "TAIL", style: { ...CHAR, hidden: true } }, // trailing zero-width run
    ];
    const { txt, caretXY } = probe(runs);
    // From just before the hidden tail, the only stops to the right are collapsed —
    // the stepper signals "cross to the next block" by returning null.
    expect(advanceVisibleStop(txt, 2, 1, false, caretXY)).toBeNull();
  });

  it("leaves ordinary text navigation unchanged (one offset per press)", () => {
    const { txt, caretXY } = probe([{ text: "hello", style: CHAR }]);
    expect(advanceVisibleStop(txt, 0, 1, false, caretXY)).toBe(1);
    expect(advanceVisibleStop(txt, 1, 1, false, caretXY)).toBe(2);
    expect(advanceVisibleStop(txt, 5, -1, false, caretXY)).toBe(4);
  });
});
