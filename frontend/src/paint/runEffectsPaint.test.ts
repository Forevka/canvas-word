// Paint-decision layer for minor run effects (issue #60): runVerticalShift folds
// the w:position raise/lower into the sub/superscript shift, widthScale turns w:w
// into a horizontal stretch factor, and doubleStrikeOffsets places the two w:dstrike
// rules. Pure math — they pin paint behavior without a canvas.
import { describe, expect, it } from "vitest";
import type { CharStyle } from "@cw/shared";
import { doubleStrikeOffsets, runVerticalShift, verticalShift, widthScale } from "./paintStyle";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#202124" };

describe("runVerticalShift — w:position raise/lower", () => {
  it("is zero with no position or sub/super", () => {
    expect(runVerticalShift(CHAR)).toBe(0);
  });

  it("raises (negative y) for a positive position, lowers for a negative one", () => {
    expect(runVerticalShift({ ...CHAR, positionPx: 4 })).toBe(-4);
    expect(runVerticalShift({ ...CHAR, positionPx: -4 })).toBe(4);
  });

  it("stacks on top of the sub/superscript shift", () => {
    const base = verticalShift("super", CHAR.fontSizePx);
    expect(runVerticalShift({ ...CHAR, verticalAlign: "super", positionPx: 2 })).toBeCloseTo(base - 2);
  });
});

describe("widthScale — w:w character scaling", () => {
  it("defaults to 1 (no scaling)", () => {
    expect(widthScale(CHAR)).toBe(1);
  });

  it("turns a percentage into a factor", () => {
    expect(widthScale({ ...CHAR, widthScalePct: 150 })).toBe(1.5);
    expect(widthScale({ ...CHAR, widthScalePct: 66 })).toBeCloseTo(0.66);
  });

  it("ignores a non-positive percentage", () => {
    expect(widthScale({ ...CHAR, widthScalePct: 0 })).toBe(1);
  });
});

describe("doubleStrikeOffsets — two w:dstrike rules", () => {
  it("returns two distinct offsets straddling the single-strike line", () => {
    const [a, b] = doubleStrikeOffsets(16);
    expect(a).toBeLessThan(b);
    expect(b - a).toBeGreaterThan(0);
  });
});
