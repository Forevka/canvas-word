// Hebrew script-split and line-breaking tests. The split routes Hebrew runs to the
// configured fallback family (matching the CJK/Arabic fallback pattern); the layout
// engine test verifies that a long Hebrew text wraps across lines (pretext line-breaks
// at space boundaries for Hebrew, matching browser behaviour).
import "./test-canvas-setup";

import { describe, it, expect, afterEach } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, Run, SectionProps } from "@cw/shared";
import { createLayoutEngine } from "./engine";
import { scriptSplitRuns, setActiveCjkFallback, setActiveArabicFallback, setActiveHebrewFallback } from "./prepareCache";

const CHAR: CharStyle = {
  fontFamily: "Arial, sans-serif",
  fontSizePx: 16,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#000000",
};
const PARA: ParaStyle = {
  align: "left",
  lineHeight: 1,
  spaceBeforePx: 0,
  spaceAfterPx: 0,
  indentFirstLinePx: 0,
  indentLeftPx: 0,
};
const SECTION: SectionProps = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
};

let nextId = 0;
const para = (runs: Run[], style: Partial<ParaStyle> = {}): Paragraph => ({
  kind: "paragraph",
  id: `t${nextId++}`,
  revision: 0,
  runs,
  style: { ...PARA, ...style },
});
const doc = (p: Paragraph): Document => ({ section: SECTION, blocks: [p] });

afterEach(() => {
  setActiveCjkFallback(null);
  setActiveArabicFallback(null);
  setActiveHebrewFallback(null);
});

describe("scriptSplitRuns (Hebrew fallback)", () => {
  it("is a no-op when no Hebrew fallback is configured", () => {
    setActiveHebrewFallback(null);
    const runs: Run[] = [{ text: "שלום עולם", style: CHAR }];
    expect(scriptSplitRuns(runs)).toBe(runs);
  });

  it("splits a mixed run into Latin + Hebrew pieces, retargeting Hebrew to the fallback", () => {
    setActiveHebrewFallback("NotoHebrew");
    const runs: Run[] = [{ text: "hello שלום world", style: CHAR }];
    const split = scriptSplitRuns(runs);
    expect(split.map((r) => r.text)).toEqual(["hello ", "שלום", " world"]);
    expect(split[0]!.style.fontFamily).toBe("Arial, sans-serif");
    expect(split[1]!.style.fontFamily).toBe("NotoHebrew");
    expect(split[2]!.style.fontFamily).toBe("Arial, sans-serif");
    // Offset-transparent: concatenation preserved.
    expect(split.map((r) => r.text).join("")).toBe("hello שלום world");
  });

  it("leaves a pure-Latin run untouched even with Hebrew fallback configured", () => {
    setActiveHebrewFallback("NotoHebrew");
    const runs: Run[] = [{ text: "hello", style: CHAR }];
    expect(scriptSplitRuns(runs)).toBe(runs);
  });

  it("handles a run that is all Hebrew", () => {
    setActiveHebrewFallback("NotoHebrew");
    const runs: Run[] = [{ text: "שלום", style: CHAR }];
    const split = scriptSplitRuns(runs);
    expect(split).toHaveLength(1);
    expect(split[0]!.style.fontFamily).toBe("NotoHebrew");
  });

  it("splits correctly when CJK, Arabic and Hebrew fallbacks are all active", () => {
    setActiveCjkFallback("NotoCJK");
    setActiveArabicFallback("NotoArabic");
    setActiveHebrewFallback("NotoHebrew");
    // "abc " Latin, "שלום" Hebrew, " " Latin, "م" Arabic, " " Latin, "中文" CJK, " xyz" Latin.
    const runs: Run[] = [{ text: "abc שלום م 中文 xyz", style: CHAR }];
    const split = scriptSplitRuns(runs);
    const texts = split.map((r) => r.text);
    const families = split.map((r) => r.style.fontFamily);
    expect(texts).toEqual(["abc ", "שלום", " ", "م", " ", "中文", " xyz"]);
    expect(families[0]).toBe("Arial, sans-serif");
    expect(families[1]).toBe("NotoHebrew");
    expect(families[3]).toBe("NotoArabic");
    expect(families[5]).toBe("NotoCJK");
    expect(texts.join("")).toBe("abc שלום م 中文 xyz");
  });

  it("does not re-split a run already targeting the Hebrew fallback family", () => {
    setActiveHebrewFallback("NotoHebrew");
    const hebrewStyle = { ...CHAR, fontFamily: "NotoHebrew" };
    const runs: Run[] = [{ text: "שלום", style: hebrewStyle }];
    expect(scriptSplitRuns(runs)).toBe(runs);
  });
});

describe("Hebrew line breaking through the engine", () => {
  it("wraps long Hebrew text at word boundaries", () => {
    const text = "שלום עולם ".repeat(20);
    const p = para([{ text, style: CHAR }], { direction: "rtl" });
    const lines = createLayoutEngine().layout(doc(p)).pages[0]!.blocks[0]!.lines;
    expect(lines.length).toBeGreaterThan(1);
    expect(lines.every((l) => l.fragments.length > 0)).toBe(true);
  });
});
