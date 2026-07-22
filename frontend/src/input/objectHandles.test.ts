import { describe, expect, it } from "vitest";
import { adaptiveHandles } from "./objectController";

// The resize handles drag via pointer-capture, which Playwright can't drive, so —
// like the rotate/crop angle math — the tiny-shape adaptation is unit-tested as a
// pure function here; the DOM wiring in `show()` just toggles display + a transform
// off this result, reusing the covered .cw-obj-handle overlay.

const names = (boxW: number, boxH: number): string[] => [...adaptiveHandles(boxW, boxH).visible].sort();
const CORNERS = ["ne", "nw", "se", "sw"]; // sorted

describe("adaptiveHandles — tiny-shape handle adaptation", () => {
  it("shows all 8 handles on a comfortably large box, no corner outset", () => {
    const a = adaptiveHandles(100, 80);
    expect(names(100, 80)).toEqual(["e", "n", "ne", "nw", "s", "se", "sw", "w"]);
    expect(a.cornerOutset).toBe(0);
  });

  it("always keeps the four corner handles", () => {
    const boxes: [number, number][] = [[100, 80], [20, 80], [80, 20], [12, 12], [1, 1]];
    for (const [w, h] of boxes) {
      for (const c of CORNERS) expect(adaptiveHandles(w, h).visible.has(c)).toBe(true);
    }
  });

  it("drops the horizontal-midpoint pair (n/s) when the box is too NARROW", () => {
    // 20px wide < 24px threshold → n and s crowd the corners and drop; e/w stay.
    const v = adaptiveHandles(20, 80).visible;
    expect(v.has("n")).toBe(false);
    expect(v.has("s")).toBe(false);
    expect(v.has("e")).toBe(true);
    expect(v.has("w")).toBe(true);
  });

  it("drops the vertical-midpoint pair (e/w) when the box is too SHORT", () => {
    const v = adaptiveHandles(100, 20).visible;
    expect(v.has("e")).toBe(false);
    expect(v.has("w")).toBe(false);
    expect(v.has("n")).toBe(true);
    expect(v.has("s")).toBe(true);
  });

  it("keeps only the corners on a box tiny on both axes", () => {
    expect(names(12, 12)).toEqual(CORNERS);
  });

  it("drops a midpoint pair exactly at the sub-threshold boundary (24px)", () => {
    // `< 24` — 24 still shows the midpoints, 23 drops them.
    expect(adaptiveHandles(24, 80).visible.has("n")).toBe(true);
    expect(adaptiveHandles(23, 80).visible.has("n")).toBe(false);
    expect(adaptiveHandles(80, 24).visible.has("e")).toBe(true);
    expect(adaptiveHandles(80, 23).visible.has("e")).toBe(false);
  });

  it("pushes corners outward only once the box is tiny on either axis (≤16px)", () => {
    expect(adaptiveHandles(100, 80).cornerOutset).toBe(0); // roomy
    expect(adaptiveHandles(17, 100).cornerOutset).toBe(0); // just above the threshold
    expect(adaptiveHandles(16, 100).cornerOutset).toBeGreaterThan(0); // at the threshold
    expect(adaptiveHandles(10, 100).cornerOutset).toBeGreaterThan(0); // one tiny axis is enough
    expect(adaptiveHandles(12, 12).cornerOutset).toBeGreaterThan(0);
  });
});
