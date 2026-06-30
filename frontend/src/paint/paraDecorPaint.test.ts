// Geometry of the paragraph-decoration box (issue #98): the shading fill and the
// w:pBdr border rectangle are expanded OUTWARD by the border-to-text padding so a
// wide/double rule sits outside the glyphs. The canvas renderer and the PDF
// painter both derive the box from this single helper, so they agree by
// construction. Pure math — pins the geometry without a canvas.
import { describe, expect, it } from "vitest";
import { paraDecorBox, PARA_BORDER_PAD_PX } from "./paintStyle";

describe("paraDecorBox — padded paragraph border/shading box", () => {
  it("expands the text box outward symmetrically by the padding", () => {
    const box = paraDecorBox(100, 200, { width: 624, height: 16, pad: PARA_BORDER_PAD_PX });
    const p = PARA_BORDER_PAD_PX;
    expect(box.x).toBe(100 - p);
    expect(box.y).toBe(200 - p);
    expect(box.width).toBe(624 + 2 * p);
    expect(box.height).toBe(16 + 2 * p);
    // The padding clears the text edges on every side.
    expect(box.x).toBeLessThan(100);
    expect(box.x + box.width).toBeGreaterThan(100 + 624);
    expect(box.y).toBeLessThan(200);
    expect(box.y + box.height).toBeGreaterThan(200 + 16);
  });

  it("is the bare text box when there is no padding (shading-only / undecorated)", () => {
    const box = paraDecorBox(100, 200, { width: 624, height: 16, pad: 0 });
    expect(box).toEqual({ x: 100, y: 200, width: 624, height: 16 });
  });

  it("treats a missing pad as zero", () => {
    const box = paraDecorBox(100, 200, { width: 624, height: 16 });
    expect(box).toEqual({ x: 100, y: 200, width: 624, height: 16 });
  });
});
