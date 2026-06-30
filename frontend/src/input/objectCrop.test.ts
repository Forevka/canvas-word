import { describe, expect, it } from "vitest";
import type { Rect } from "../layout/geometry";
import { fullSourceRect, windowRectFromInsets, cropInsetsFromWindow, isEmptyCrop, type CropInsets } from "./objectController";

const rect = (x: number, y: number, width: number, height: number): Rect => ({ x, y, width, height, pageIndex: 0 });
const near = (a: number, b: number): void => expect(a).toBeCloseTo(b, 6);

describe("crop overlay geometry", () => {
  it("an un-cropped image's full source rect IS its displayed box", () => {
    const r = rect(100, 50, 200, 120);
    const full = fullSourceRect(r, { left: 0, top: 0, right: 0, bottom: 0 });
    expect(full).toMatchObject({ x: 100, y: 50, width: 200, height: 120 });
  });

  it("derives the full source rect so the current crop window maps onto the box", () => {
    // 200px box shows the middle 50% horizontally (left .25 / right .25) → full is 400px,
    // shifted left by .25*400 = 100 so the window [.25,.75] lands back on the box.
    const r = rect(100, 50, 200, 100);
    const crop: CropInsets = { left: 0.25, top: 0.2, right: 0.25, bottom: 0.2 };
    const full = fullSourceRect(r, crop);
    near(full.width, 400);
    near(full.height, 100 / 0.6);
    near(full.x, 100 - 0.25 * 400);
    // The window rect back out of that full rect must equal the original displayed box.
    const win = windowRectFromInsets(full, crop);
    near(win.x, r.x);
    near(win.y, r.y);
    near(win.width, r.width);
    near(win.height, r.height);
  });

  it("round-trips insets ⇄ window within the full source rect", () => {
    const full = rect(0, 0, 400, 300);
    const crop: CropInsets = { left: 0.1, top: 0.2, right: 0.05, bottom: 0.15 };
    const win = windowRectFromInsets(full, crop);
    const back = cropInsetsFromWindow(full, win);
    near(back.left, crop.left);
    near(back.top, crop.top);
    near(back.right, crop.right);
    near(back.bottom, crop.bottom);
  });

  it("a window dragged to the top-left quadrant yields right=bottom=0.5 insets", () => {
    const full = rect(0, 0, 400, 200);
    const insets = cropInsetsFromWindow(full, rect(0, 0, 200, 100));
    expect(insets).toEqual({ left: 0, top: 0, right: 0.5, bottom: 0.5 });
  });

  it("clamps insets into [0,1) even when the window pokes outside the source", () => {
    const full = rect(0, 0, 100, 100);
    const insets = cropInsetsFromWindow(full, rect(-20, -20, 200, 200));
    expect(insets.left).toBe(0);
    expect(insets.top).toBe(0);
    expect(insets.right).toBe(0);
    expect(insets.bottom).toBe(0);
  });

  it("isEmptyCrop only treats a (near) full-frame window as no crop", () => {
    expect(isEmptyCrop({ left: 0, top: 0, right: 0, bottom: 0 })).toBe(true);
    expect(isEmptyCrop({ left: 0.0001, top: 0, right: 0, bottom: 0 })).toBe(true);
    expect(isEmptyCrop({ left: 0.05, top: 0, right: 0, bottom: 0 })).toBe(false);
  });
});
