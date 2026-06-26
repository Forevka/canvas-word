// Mathematical Alphanumeric remapping — the table that lets a single upright math
// font render true italic/bold/blackboard/script glyphs.

import { describe, expect, it } from "vitest";
import { styledMathText } from "./mathAlpha";

describe("styledMathText", () => {
  it("maps ASCII letters to math-italic by default", () => {
    expect(styledMathText("x", undefined)).toBe("\u{1D465}"); // 𝑥
    expect(styledMathText("A", undefined)).toBe("\u{1D434}"); // 𝐴
  });

  it("applies the italic 'h' Planck-constant exception", () => {
    expect(styledMathText("h", undefined)).toBe("ℎ"); // ℎ
  });

  it("maps bold and bold-italic", () => {
    expect(styledMathText("A", "bold")).toBe("\u{1D400}");
    expect(styledMathText("a", "bold-italic")).toBe("\u{1D482}");
  });

  it("uses Letterlike holes for blackboard letters", () => {
    expect(styledMathText("R", "double-struck")).toBe("ℝ"); // ℝ
    expect(styledMathText("D", "double-struck")).toBe("\u{1D53B}"); // contiguous
  });

  it("keeps normal-variant ASCII upright", () => {
    expect(styledMathText("x", "normal")).toBe("x");
  });

  it("leaves multi-character identifiers (function names) unchanged", () => {
    expect(styledMathText("sin", undefined)).toBe("sin");
    expect(styledMathText("log", "bold")).toBe("log");
  });

  it("maps Greek to math-italic Greek", () => {
    expect(styledMathText("α", undefined)).toBe("\u{1D6FC}"); // 𝛼
  });

  it("passes operators and digits through", () => {
    expect(styledMathText("+", undefined)).toBe("+");
    expect(styledMathText("2", undefined)).toBe("2");
  });
});
