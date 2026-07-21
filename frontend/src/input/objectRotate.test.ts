import { describe, expect, it } from "vitest";
import type { Document, ImageBlock } from "@cw/shared";
import { applyOp } from "@cw/shared";
import { normalizeDeg, snapRotation, pointerAngleDeg, rotationFromDrag } from "./objectController";
import { ICONS } from "../ui/icons";

// The rotate handle drives a pointer-capture drag the same way the resize/crop
// handles do — and, like them, Playwright can't drive pointer-capture drags. So we
// unit-test the pure angle math + snapping and the rotate *command* (the model op the
// handle commits) here; the handle's DOM wiring reuses the covered .cw-obj-handle
// pointer-capture pattern. (No jsdom env is configured, so the frame's DOM render is
// verified structurally via the icon glyph rather than a live mount.)

describe("rotate-handle angle math", () => {
  it("normalizeDeg wraps into [0,360)", () => {
    expect(normalizeDeg(0)).toBe(0);
    expect(normalizeDeg(370)).toBe(10);
    expect(normalizeDeg(-30)).toBe(330);
    expect(normalizeDeg(360)).toBe(0);
    expect(normalizeDeg(720 + 45)).toBe(45);
  });

  it("snapRotation snaps to the nearest 15° only when held", () => {
    expect(snapRotation(22, false)).toBe(22); // free drag — unchanged
    expect(snapRotation(22, true)).toBe(15); // snaps down
    expect(snapRotation(23, true)).toBe(30); // snaps up
    expect(snapRotation(7, true)).toBe(0); // nearest grid line
    expect(snapRotation(52, true, 45)).toBe(45); // custom step
  });

  it("pointerAngleDeg is a y-down clockwise atan2 about the center", () => {
    expect(pointerAngleDeg(0, 0, 10, 0)).toBe(0); // due right
    expect(pointerAngleDeg(0, 0, 0, 10)).toBe(90); // straight down (y-down)
    expect(pointerAngleDeg(0, 0, -10, 0)).toBe(180); // due left
    expect(pointerAngleDeg(0, 0, 0, -10)).toBe(-90); // straight up
  });

  it("rotationFromDrag adds the swept pointer angle to the start rotation", () => {
    // Pointer sweeps 0° → 30° with no starting rotation → 30°.
    expect(rotationFromDrag(0, 0, 30, false)).toBe(30);
    // Starting at 40°, pointer sweeps 10° → 90° (+80°) → 120°.
    expect(rotationFromDrag(40, 10, 90, false)).toBe(120);
    // Sweeping backwards past zero normalizes into [0,360).
    expect(rotationFromDrag(10, 0, -30, false)).toBe(340);
  });

  it("rotationFromDrag snaps the result to 15° when Shift is held", () => {
    expect(rotationFromDrag(0, 0, 22, true)).toBe(15);
    expect(rotationFromDrag(0, 0, 38, true)).toBe(45);
    // Snap applies to the absolute result, not the delta.
    expect(rotationFromDrag(7, 0, 5, true)).toBe(15); // 12 → nearest 15
  });
});

// The rotate command the handle commits: setImageProps { rotation } (mirrors the
// pre-existing shape rotation reducer, covered by shapesRoundtrip).
const SECTION = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };
const img = (rotation?: number): ImageBlock => ({
  kind: "image", id: "i", revision: 0, src: "blob:x", widthPx: 100, heightPx: 80, align: "center",
  ...(rotation !== undefined ? { rotation } : {}),
});
const docOf = (b: ImageBlock): Document => ({ section: SECTION, blocks: [b] });
const imageOf = (d: Document): ImageBlock => d.blocks[0] as ImageBlock;

describe("setImageProps rotation (the rotate command)", () => {
  it("sets a rotation on an unrotated image", () => {
    const res = applyOp(docOf(img()), { type: "setImageProps", blockId: "i", patch: { rotation: 18 } });
    expect(imageOf(res.doc).rotation).toBe(18);
  });

  it("null clears the rotation", () => {
    const res = applyOp(docOf(img(30)), { type: "setImageProps", blockId: "i", patch: { rotation: null } });
    expect(imageOf(res.doc).rotation).toBeUndefined();
  });

  it("is undoable back to the prior rotation", () => {
    const set = applyOp(docOf(img(15)), { type: "setImageProps", blockId: "i", patch: { rotation: 90 } });
    expect(imageOf(set.doc).rotation).toBe(90);
    const undone = applyOp(set.doc, set.inverse).doc;
    expect(imageOf(undone).rotation).toBe(15); // inverse restores the old angle
  });

  it("undo of a set-from-none clears the field again", () => {
    const set = applyOp(docOf(img()), { type: "setImageProps", blockId: "i", patch: { rotation: 45 } });
    const undone = applyOp(set.doc, set.inverse).doc;
    expect(imageOf(undone).rotation).toBeUndefined();
  });
});

describe("rotate handle glyph", () => {
  it("ICONS.rotate is a well-formed single-color svg (renders in the handle)", () => {
    expect(ICONS.rotate).toMatch(/^<svg/);
    expect(ICONS.rotate).toContain('viewBox="0 0 16 16"');
    expect(ICONS.rotate).toContain("stroke=\"currentColor\"");
    expect(ICONS.rotate).toContain("<path");
  });
});
