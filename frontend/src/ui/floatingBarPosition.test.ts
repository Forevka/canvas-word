import { describe, expect, it } from "vitest";
import { anchorInView, placeSelectionBar } from "./floatingBarPosition";

const VP = { width: 1000, height: 800 };

describe("placeSelectionBar", () => {
  it("centers the bar horizontally over the anchor and sits above it", () => {
    // Anchor at x=400..500 (center 450), bar 200 wide → left = 450 - 100 = 350.
    const anchor = { left: 400, top: 300, width: 100, height: 20 };
    const { left, top } = placeSelectionBar(anchor, { width: 200, height: 30 }, VP);
    expect(left).toBe(350);
    // top = anchor.top - barH - gap(8) = 300 - 30 - 8 = 262.
    expect(top).toBe(262);
  });

  it("flips below the anchor when there is no room above the top guard", () => {
    const anchor = { left: 400, top: 40, width: 100, height: 20 };
    const { top } = placeSelectionBar(anchor, { width: 200, height: 30 }, VP);
    // above = 40 - 30 - 8 = 2 < topGuard(56) → below = top + height + gap = 40 + 20 + 8.
    expect(top).toBe(68);
  });

  it("clamps to the left margin when the anchor is near the left edge", () => {
    const anchor = { left: 0, top: 400, width: 20, height: 20 };
    const { left } = placeSelectionBar(anchor, { width: 200, height: 30 }, VP);
    expect(left).toBe(8); // default margin
  });

  it("clamps to the right margin when the anchor is near the right edge", () => {
    const anchor = { left: 980, top: 400, width: 20, height: 20 };
    const { left } = placeSelectionBar(anchor, { width: 200, height: 30 }, VP);
    // maxLeft = viewport.width - barW - margin = 1000 - 200 - 8 = 792.
    expect(left).toBe(792);
  });

  it("honors a custom gap", () => {
    const anchor = { left: 400, top: 300, width: 100, height: 20 };
    const { top } = placeSelectionBar(anchor, { width: 200, height: 30 }, VP, { gap: 20 });
    expect(top).toBe(250); // 300 - 30 - 20
  });

  it("uses viewport.top as the guard so it never renders over the ribbon", () => {
    // Anchor just below a 140px ribbon: room "above" (top 100) would land in the
    // ribbon, so it must flip below instead of placing at 100.
    const anchor = { left: 400, top: 150, width: 100, height: 20 };
    const { top } = placeSelectionBar(anchor, { width: 200, height: 30 }, { ...VP, top: 140 });
    // above = 150 - 30 - 8 = 112 < topGuard(140) → below = 150 + 20 + 8 = 178.
    expect(top).toBe(178);
  });

  it("never places above viewport.top even when flipping below would go higher", () => {
    const anchor = { left: 400, top: 130, width: 100, height: 4 };
    const { top } = placeSelectionBar(anchor, { width: 200, height: 30 }, { ...VP, top: 140 });
    expect(top).toBeGreaterThanOrEqual(140);
  });
});

describe("anchorInView", () => {
  it("is true for an anchor inside the viewport", () => {
    expect(anchorInView({ left: 100, top: 100, width: 50, height: 20 }, VP)).toBe(true);
  });

  it("is false for an anchor scrolled above the viewport", () => {
    expect(anchorInView({ left: 100, top: -50, width: 50, height: 20 }, VP)).toBe(false);
  });

  it("is false for an anchor scrolled below the viewport", () => {
    expect(anchorInView({ left: 100, top: 900, width: 50, height: 20 }, VP)).toBe(false);
  });

  it("is false for an anchor scrolled up behind the ribbon (above viewport.top)", () => {
    // Bottom at 120 is above the 140px content top → hidden, not left over the ribbon.
    expect(anchorInView({ left: 100, top: 100, width: 50, height: 20 }, { ...VP, top: 140 })).toBe(false);
    // Still visible once its bottom crosses into the content area.
    expect(anchorInView({ left: 100, top: 135, width: 50, height: 20 }, { ...VP, top: 140 })).toBe(true);
  });
});
