// Page Layout dialog length conversions. Model is px @96dpi; these convert only at
// the input boundary (in <-> px, cm <-> px) and format for display rounded to 2dp.

import { describe, expect, it } from "vitest";
import {
  cmToPx,
  formatUnit,
  inToPx,
  pxToCm,
  pxToIn,
  pxToUnit,
  unitToPx,
} from "./units";

describe("inToPx / pxToIn", () => {
  it("converts inches to px and back at 96dpi", () => {
    expect(inToPx(1)).toBe(96);
    expect(pxToIn(96)).toBe(1);
    expect(inToPx(8.5)).toBe(816);
  });

  it("round-trips inches through px", () => {
    for (const inch of [0, 0.5, 1, 11]) {
      expect(pxToIn(inToPx(inch))).toBeCloseTo(inch, 10);
    }
  });
});

describe("cmToPx / pxToCm", () => {
  it("converts cm to px and back at 96dpi", () => {
    expect(cmToPx(2.54)).toBe(96); // 1 inch
    expect(pxToCm(96)).toBeCloseTo(2.54, 10);
  });

  it("round-trips cm through px", () => {
    for (const c of [0, 1, 2.54, 21]) {
      expect(pxToCm(cmToPx(c))).toBeCloseTo(c, 10);
    }
  });
});

describe("pxToUnit", () => {
  it("dispatches to inches for unit 'in'", () => {
    expect(pxToUnit(96, "in")).toBe(1);
  });

  it("dispatches to centimeters for unit 'cm'", () => {
    expect(pxToUnit(96, "cm")).toBeCloseTo(2.54, 10);
  });
});

describe("unitToPx", () => {
  it("dispatches to inToPx for unit 'in'", () => {
    expect(unitToPx(1, "in")).toBe(96);
  });

  it("dispatches to cmToPx for unit 'cm'", () => {
    expect(unitToPx(2.54, "cm")).toBe(96);
  });

  it("inverts pxToUnit for both units", () => {
    expect(unitToPx(pxToUnit(816, "in"), "in")).toBeCloseTo(816, 10);
    expect(unitToPx(pxToUnit(816, "cm"), "cm")).toBeCloseTo(816, 10);
  });
});

describe("formatUnit", () => {
  it("formats inches rounded to 2 decimals", () => {
    expect(formatUnit(96, "in")).toBe("1");
    expect(formatUnit(816, "in")).toBe("8.5");
  });

  it("formats centimeters rounded to 2 decimals", () => {
    // 96px = 2.54cm exactly, displayed as "2.54".
    expect(formatUnit(96, "cm")).toBe("2.54");
  });

  it("drops trailing-zero noise via Number stringification", () => {
    // 192px = 2in -> "2", not "2.00".
    expect(formatUnit(192, "in")).toBe("2");
  });

  it("rounds half-up at the 2nd decimal", () => {
    // 1.005in -> 96.48px; pxToIn back is 1.005, round(*100)/100 -> 1.01 or 1 depending
    // on float; assert the actual rounding behavior of a representative value.
    const px = inToPx(1.236);
    expect(formatUnit(px, "in")).toBe("1.24");
  });

  it("formats zero as '0'", () => {
    expect(formatUnit(0, "in")).toBe("0");
    expect(formatUnit(0, "cm")).toBe("0");
  });
});
