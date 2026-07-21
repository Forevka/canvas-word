import { describe, expect, it } from "vitest";
import { bandRegionRect, type BandPageGeom } from "./bandHoverController";

const PAGE: BandPageGeom = {
  index: 2,
  widthPx: 816, // 8.5in @96dpi
  heightPx: 1056, // 11in @96dpi
  marginPx: { left: 96, right: 96 },
  contentTopPx: 96,
  contentBottomPx: 960,
};

describe("bandRegionRect", () => {
  it("spans the content width for both bands", () => {
    const contentWidth = PAGE.widthPx - PAGE.marginPx.left - PAGE.marginPx.right; // 624
    expect(bandRegionRect(PAGE, "header").x).toBe(PAGE.marginPx.left);
    expect(bandRegionRect(PAGE, "header").width).toBe(contentWidth);
    expect(bandRegionRect(PAGE, "footer").x).toBe(PAGE.marginPx.left);
    expect(bandRegionRect(PAGE, "footer").width).toBe(contentWidth);
  });

  it("puts the header from the page top down to the body content top", () => {
    expect(bandRegionRect(PAGE, "header")).toEqual({
      pageIndex: 2,
      x: 96,
      y: 0,
      width: 624,
      height: 96, // contentTopPx
    });
  });

  it("puts the footer from the body content bottom to the page bottom", () => {
    expect(bandRegionRect(PAGE, "footer")).toEqual({
      pageIndex: 2,
      x: 96,
      y: 960, // contentBottomPx
      width: 624,
      height: 96, // heightPx - contentBottomPx
    });
  });

  it("tracks a tall header that has pushed the content box down", () => {
    // A header taller than the top margin lowers contentTopPx's counterpart —
    // here contentTopPx grew to 200, so the header region grows with it.
    const tall: BandPageGeom = { ...PAGE, contentTopPx: 200 };
    expect(bandRegionRect(tall, "header").height).toBe(200);
  });

  it("carries the page index through so the overlay mounts on the right page", () => {
    expect(bandRegionRect({ ...PAGE, index: 7 }, "footer").pageIndex).toBe(7);
  });
});
