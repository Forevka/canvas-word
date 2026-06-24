// Builder unit helpers convert physical units to model px @96dpi at the API
// boundary, and PAGE_SIZES carries the editor's standard page geometries in px.

import { describe, expect, it } from "vitest";
import { cm, inches, PAGE_SIZES, pt, twips } from "./units";

describe("inches", () => {
  it("converts inches to px at 96dpi", () => {
    expect(inches(1)).toBe(96);
    expect(inches(8.5)).toBe(816); // Letter width
    expect(inches(0)).toBe(0);
  });
});

describe("cm", () => {
  it("converts centimeters to px at 96dpi", () => {
    expect(cm(2.54)).toBe(96); // 1 inch
    expect(cm(0)).toBe(0);
    expect(cm(21)).toBeCloseTo(793.7, 1); // A4 width
  });
});

describe("pt", () => {
  it("converts points (1/72in) to px at 96dpi", () => {
    expect(pt(72)).toBe(96); // 1 inch
    expect(pt(12)).toBe(16); // 12pt body text
    expect(pt(0)).toBe(0);
  });
});

describe("twips", () => {
  it("converts twips (1/20pt) to px, matching the import factor of /15", () => {
    expect(twips(1440)).toBe(96); // 1 inch
    expect(twips(15)).toBe(1);
    expect(twips(240)).toBe(16); // 12pt
  });
});

describe("unit helper relationships", () => {
  it("inches and pt agree (72pt = 1in)", () => {
    expect(pt(72)).toBe(inches(1));
  });

  it("twips and pt agree (20twip = 1pt)", () => {
    expect(twips(20)).toBe(pt(1));
  });
});

describe("PAGE_SIZES", () => {
  it("exposes Letter as the editor default (816x1056 px)", () => {
    expect(PAGE_SIZES.Letter).toEqual({ pageWidthPx: 816, pageHeightPx: 1056 });
  });

  it("derives Letter dimensions from inches (8.5 x 11)", () => {
    expect(PAGE_SIZES.Letter.pageWidthPx).toBe(inches(8.5));
    expect(PAGE_SIZES.Letter.pageHeightPx).toBe(inches(11));
  });

  it("includes the common named sizes", () => {
    expect(Object.keys(PAGE_SIZES).sort()).toEqual(
      ["A3", "A4", "Legal", "Letter", "Tabloid"].sort(),
    );
  });

  it("gives every page a positive width and height", () => {
    for (const size of Object.values(PAGE_SIZES)) {
      expect(size.pageWidthPx).toBeGreaterThan(0);
      expect(size.pageHeightPx).toBeGreaterThan(0);
      expect(size.pageHeightPx).toBeGreaterThanOrEqual(size.pageWidthPx); // portrait
    }
  });
});
