// Arabic script-split and line-breaking tests. The split routes Arabic runs to the
// configured fallback family (matching the CJK fallback pattern); the layout engine
// test verifies that a long Arabic text wraps across lines (pretext line-breaks at
// space boundaries for Arabic, matching browser behaviour).
import "./test-canvas-setup";

import { describe, it, expect, afterEach } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, Run, SectionProps } from "@cw/shared";
import { createLayoutEngine } from "./engine";
import { scriptSplitRuns, setActiveCjkFallback, setActiveArabicFallback } from "./prepareCache";

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
});

describe("scriptSplitRuns (Arabic fallback)", () => {
  it("is a no-op when no Arabic fallback is configured", () => {
    setActiveArabicFallback(null);
    const runs: Run[] = [{ text: "مرحبا بالعالم", style: CHAR }];
    expect(scriptSplitRuns(runs)).toBe(runs);
  });

  it("splits a mixed run into Latin + Arabic pieces, retargeting Arabic to the fallback", () => {
    setActiveArabicFallback("NotoArabic");
    const runs: Run[] = [{ text: "hello مرحبا world", style: CHAR }];
    const split = scriptSplitRuns(runs);
    expect(split.map((r) => r.text)).toEqual(["hello ", "مرحبا", " world"]);
    expect(split[0]!.style.fontFamily).toBe("Arial, sans-serif");
    expect(split[1]!.style.fontFamily).toBe("NotoArabic");
    expect(split[2]!.style.fontFamily).toBe("Arial, sans-serif");
    // Offset-transparent: concatenation preserved.
    expect(split.map((r) => r.text).join("")).toBe("hello مرحبا world");
  });

  it("leaves a pure-Latin run untouched even with Arabic fallback configured", () => {
    setActiveArabicFallback("NotoArabic");
    const runs: Run[] = [{ text: "hello", style: CHAR }];
    expect(scriptSplitRuns(runs)).toBe(runs);
  });

  it("handles a run that is all Arabic", () => {
    setActiveArabicFallback("NotoArabic");
    const runs: Run[] = [{ text: "مرحبا", style: CHAR }];
    const split = scriptSplitRuns(runs);
    expect(split).toHaveLength(1);
    expect(split[0]!.style.fontFamily).toBe("NotoArabic");
  });

  it("splits correctly when both CJK and Arabic fallbacks are active", () => {
    setActiveCjkFallback("NotoCJK");
    setActiveArabicFallback("NotoArabic");
    const runs: Run[] = [{ text: "abc مرحبا 中文 xyz", style: CHAR }];
    const split = scriptSplitRuns(runs);
    const texts = split.map((r) => r.text);
    const families = split.map((r) => r.style.fontFamily);
    // "abc " is Latin, "مرحبا" Arabic, " " Latin, "中文" CJK, " xyz" Latin
    expect(texts).toEqual(["abc ", "مرحبا", " ", "中文", " xyz"]);
    expect(families[0]).toBe("Arial, sans-serif");
    expect(families[1]).toBe("NotoArabic");
    expect(families[2]).toBe("Arial, sans-serif");
    expect(families[3]).toBe("NotoCJK");
    expect(families[4]).toBe("Arial, sans-serif");
    expect(texts.join("")).toBe("abc مرحبا 中文 xyz");
  });

  it("does not re-split a run already targeting the Arabic fallback family", () => {
    setActiveArabicFallback("NotoArabic");
    const arabicStyle = { ...CHAR, fontFamily: "NotoArabic" };
    const runs: Run[] = [{ text: "مرحبا", style: arabicStyle }];
    expect(scriptSplitRuns(runs)).toBe(runs);
  });
});

describe("Arabic line breaking through the engine", () => {
  it("wraps long Arabic text at word boundaries", () => {
    // A paragraph of repeated Arabic words wide enough to require line breaks.
    const text = "مرحبا بالعالم ".repeat(20);
    const p = para([{ text, style: CHAR }], { direction: "rtl" });
    const lines = createLayoutEngine().layout(doc(p)).pages[0]!.blocks[0]!.lines;
    expect(lines.length).toBeGreaterThan(1);
    expect(lines.every((l) => l.fragments.length > 0)).toBe(true);
  });
});
