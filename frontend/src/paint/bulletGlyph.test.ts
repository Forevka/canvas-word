// Issue #99: the three default bullet levels (•/◦/▪) are drawn as VECTOR SHAPES
// rather than glyphs so ◦ (U+25E6) and ▪ (U+25AA) — absent from the bundled PDF
// Latin font subset — never render as .notdef/tofu in the export. bulletShapeFor is
// the shared decision both painters use; these tests pin the shape selection and
// geometry without a canvas or PDFKit.
import { describe, expect, it } from "vitest";
import { BULLET_DISC, BULLET_RING, BULLET_SQUARE, bulletShapeFor } from "./paintStyle";

const FS = 16;
const X = 100;
const BASELINE = 200;

describe("bulletShapeFor — default bullet → vector shape", () => {
  it("maps • (U+2022) to a filled disc", () => {
    const shape = bulletShapeFor(BULLET_DISC, FS, X, BASELINE);
    expect(shape?.kind).toBe("disc");
    if (shape?.kind !== "disc") throw new Error("expected disc");
    expect(shape.r).toBeGreaterThan(0);
    // Centred a touch right of the marker's left edge and above the baseline.
    expect(shape.cx).toBeGreaterThan(X);
    expect(shape.cy).toBeLessThan(BASELINE);
  });

  it("maps ◦ (U+25E6 WHITE BULLET) to a stroked ring", () => {
    const shape = bulletShapeFor(BULLET_RING, FS, X, BASELINE);
    expect(shape?.kind).toBe("ring");
    if (shape?.kind !== "ring") throw new Error("expected ring");
    expect(shape.r).toBeGreaterThan(0);
    expect(shape.lineWidth).toBeGreaterThan(0);
    // The ring's stroke must be thinner than its radius (otherwise it fills in).
    expect(shape.lineWidth).toBeLessThan(shape.r);
  });

  it("maps ▪ (U+25AA BLACK SMALL SQUARE) to a filled square", () => {
    const shape = bulletShapeFor(BULLET_SQUARE, FS, X, BASELINE);
    expect(shape?.kind).toBe("square");
    if (shape?.kind !== "square") throw new Error("expected square");
    expect(shape.size).toBeGreaterThan(0);
    // The box is centred on the bullet centre (above the baseline, right of x).
    expect(shape.x).toBeGreaterThan(X);
    expect(shape.y + shape.size).toBeLessThan(BASELINE);
  });

  it("scales the shape with the font size", () => {
    const small = bulletShapeFor(BULLET_DISC, 10, X, BASELINE);
    const large = bulletShapeFor(BULLET_DISC, 30, X, BASELINE);
    if (small?.kind !== "disc" || large?.kind !== "disc") throw new Error("expected discs");
    expect(large.r).toBeGreaterThan(small.r);
  });

  it("returns undefined for numbered and custom markers (they paint as text)", () => {
    expect(bulletShapeFor("1.", FS, X, BASELINE)).toBeUndefined();
    expect(bulletShapeFor("a)", FS, X, BASELINE)).toBeUndefined();
    expect(bulletShapeFor("›", FS, X, BASELINE)).toBeUndefined(); // a custom bullet glyph
    expect(bulletShapeFor("", FS, X, BASELINE)).toBeUndefined();
  });
});
