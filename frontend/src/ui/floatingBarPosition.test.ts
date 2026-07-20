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
});
