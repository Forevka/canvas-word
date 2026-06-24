// px @96dpi -> OOXML units. These are the exact inverses of import/docx/units.ts;
// the round-trip relationships (px -> tw -> px = identity for whole inches) are the
// load-bearing invariants, plus the non-finite coercion and the eighth-point clamp.

import { describe, expect, it } from "vitest";
import {
  multiplierToLine,
  pxToEighthPoints,
  pxToEmu,
  pxToHalfPoints,
  pxToTwips,
} from "./units";

describe("pxToTwips", () => {
  it("converts px to twips at 15 twips/px", () => {
    expect(pxToTwips(96)).toBe(1440); // 1 inch
    expect(pxToTwips(48)).toBe(720); // 0.5 inch
    expect(pxToTwips(16)).toBe(240); // 12pt
  });

  it("rounds to the nearest integer twip", () => {
    expect(pxToTwips(1)).toBe(15);
    expect(pxToTwips(1.03)).toBe(15); // 15.45 -> 15
    expect(pxToTwips(1.04)).toBe(16); // 15.6 -> 16
  });

  it("coerces a non-finite value to 0", () => {
    expect(pxToTwips(NaN)).toBe(0);
    expect(pxToTwips(Infinity)).toBe(0);
    expect(pxToTwips(-Infinity)).toBe(0);
  });

  it("handles zero and negatives", () => {
    expect(pxToTwips(0)).toBe(0);
    expect(pxToTwips(-16)).toBe(-240);
  });
});

describe("pxToHalfPoints", () => {
  it("converts px to half-points at 1.5 hp/px", () => {
    expect(pxToHalfPoints(16)).toBe(24); // 12pt font
    expect(pxToHalfPoints(8)).toBe(12);
  });

  it("rounds to the nearest half-point", () => {
    expect(pxToHalfPoints(14.6667)).toBe(22); // 22.0 -> 22 (Word's 11pt default)
  });

  it("coerces non-finite to 0", () => {
    expect(pxToHalfPoints(NaN)).toBe(0);
  });
});

describe("pxToEmu", () => {
  it("converts px to EMU at 9525 EMU/px", () => {
    expect(pxToEmu(96)).toBe(914400); // 1 inch
    expect(pxToEmu(1)).toBe(9525);
  });

  it("coerces non-finite to 0", () => {
    expect(pxToEmu(Infinity)).toBe(0);
  });
});

describe("pxToEighthPoints", () => {
  it("converts px to eighth-points at 6 sz/px", () => {
    expect(pxToEighthPoints(1)).toBe(6);
    expect(pxToEighthPoints(2)).toBe(12);
  });

  it("clamps the result up to a minimum of 2 (hairline stays visible)", () => {
    expect(pxToEighthPoints(0)).toBe(2);
    expect(pxToEighthPoints(0.1)).toBe(2); // round(0.6)=1 -> clamped to 2
    expect(pxToEighthPoints(0.33)).toBe(2); // round(1.98)=2
  });

  it("coerces non-finite to the clamped minimum", () => {
    // finite(NaN)=0 -> round(0)=0 -> max(2,0)=2
    expect(pxToEighthPoints(NaN)).toBe(2);
  });
});

describe("multiplierToLine", () => {
  it("converts a line-height multiplier to 240ths", () => {
    expect(multiplierToLine(1)).toBe(240);
    expect(multiplierToLine(1.5)).toBe(360);
    expect(multiplierToLine(2)).toBe(480);
  });

  it("coerces non-finite to 0", () => {
    expect(multiplierToLine(NaN)).toBe(0);
  });
});

describe("px <-> OOXML round-trip invariants", () => {
  it("twips invert import twipsToPx for whole inches", () => {
    // 1 inch = 96px = 1440tw; pxToTwips is the documented inverse of twipsToPx.
    for (const inch of [0.5, 1, 2, 8.5]) {
      const px = inch * 96;
      expect(pxToTwips(px) / 15).toBe(px);
    }
  });

  it("EMU invert import emuToPx exactly (no rounding loss at whole px)", () => {
    for (const px of [1, 10, 96, 816]) {
      expect(pxToEmu(px) / 9525).toBe(px);
    }
  });
});
