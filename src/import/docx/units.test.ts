import { describe, expect, it } from "vitest";
import { emuToPx, halfPointsToPx, lineAutoToMultiplier, round2, twipsToPx } from "./units";

describe("units", () => {
  it("converts twips to px (96dpi)", () => {
    expect(twipsToPx(1440)).toBe(96); // 1 inch
    expect(twipsToPx(720)).toBe(48); // 0.5 inch
    expect(twipsToPx(240)).toBe(16); // 12pt
  });

  it("converts half-points to px", () => {
    expect(halfPointsToPx(24)).toBe(16); // 12pt
    expect(halfPointsToPx(22)).toBeCloseTo(14.6667, 3); // 11pt — Word's default
  });

  it("converts EMU to px", () => {
    expect(emuToPx(914400)).toBe(96); // 1 inch
  });

  it("converts auto line spacing to a multiplier", () => {
    expect(lineAutoToMultiplier(240)).toBe(1);
    expect(lineAutoToMultiplier(360)).toBe(1.5);
    expect(lineAutoToMultiplier(480)).toBe(2);
  });

  it("rounds to 2 decimals", () => {
    expect(round2(14.66666)).toBe(14.67);
    expect(round2(96)).toBe(96);
  });
});
