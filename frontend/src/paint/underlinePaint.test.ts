// Paint-decision layer for underline styles: runPaint() resolves the line style +
// colour both painters consume, and underlinePlan()/underlineWavePoints() turn a
// style into the concrete strokes the canvas renderer and PDF painter share. These
// are the metric-neutral "what does the rule look like" decisions — pure, so they
// pin double/dotted/dashed/wave without a canvas.
import { describe, expect, it } from "vitest";
import type { CharStyle } from "@cw/shared";
import { runPaint, underlinePlan, underlineWavePoints } from "./paintStyle";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#202124" };

describe("runPaint — underline style + colour", () => {
  it("carries the run's underline style and follows the text colour by default", () => {
    const rp = runPaint({ ...CHAR, underline: true, underlineStyle: "double" });
    expect(rp.underline).toBe(true);
    expect(rp.underlineStyle).toBe("double");
    expect(rp.underlineColor).toBe("#202124"); // no explicit colour → text colour
  });

  it("uses an explicit underline colour when set", () => {
    const rp = runPaint({ ...CHAR, underline: true, underlineStyle: "wave", underlineColor: "#d93025" });
    expect(rp.underlineColor).toBe("#d93025");
  });

  it("defaults a styleless underline to a plain single line", () => {
    expect(runPaint({ ...CHAR, underline: true }).underlineStyle).toBe("single");
  });

  it("draws external-link affordance as a plain single rule in the link colour", () => {
    const rp = runPaint({ ...CHAR, link: "https://example.com", underline: true, underlineStyle: "wave", underlineColor: "#d93025" }, "#0b57d0");
    expect(rp.underline).toBe(true);
    expect(rp.underlineStyle).toBe("single"); // links never inherit a fancy rule
    expect(rp.underlineColor).toBe("#0b57d0"); // and paint in the link colour
  });
});

describe("underlinePlan — style → strokes", () => {
  it("double draws two parallel solid lines", () => {
    const p = underlinePlan("double", 16);
    expect(p.double).toBe(true);
    expect(p.wave).toBe(false);
    expect(p.dash).toEqual([]);
  });

  it("dashed and dotted produce a dash pattern", () => {
    expect(underlinePlan("dash", 16).dash.length).toBeGreaterThan(0);
    expect(underlinePlan("dotted", 16).dash.length).toBeGreaterThan(0);
  });

  it("dotDash / dotDotDash alternate dashes and dots", () => {
    expect(underlinePlan("dotDash", 16).dash.length).toBe(4);
    expect(underlinePlan("dotDotDash", 16).dash.length).toBe(6);
  });

  it("thick is at least 2px and noticeably heavier than single", () => {
    expect(underlinePlan("thick", 16).thickness).toBeGreaterThanOrEqual(2);
    expect(underlinePlan("thick", 16).thickness).toBeGreaterThan(underlinePlan("single", 16).thickness);
  });

  it("wave flags a sine stroke whose sampled points oscillate", () => {
    const p = underlinePlan("wave", 16);
    expect(p.wave).toBe(true);
    const pts = underlineWavePoints(0, 10, 40, 16);
    const ys = pts.map((q) => q.y);
    expect(Math.max(...ys)).toBeGreaterThan(10); // crest above the centreline
    expect(Math.min(...ys)).toBeLessThan(10); // trough below it
  });
});
