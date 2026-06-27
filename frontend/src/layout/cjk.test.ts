// CJK line-breaking + font-fallback tests. Breaking/kinsoku come from pretext
// (no config needed); these assert the run-splitting fallback and that mixed
// CJK/Latin paragraphs break between CJK characters.
import "./test-canvas-setup";

import { describe, it, expect, afterEach } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, Run, SectionProps } from "@cw/shared";
import { createLayoutEngine } from "./engine";
import { scriptSplitRuns, setActiveCjkFallback } from "./prepareCache";

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

afterEach(() => setActiveCjkFallback(null));

describe("scriptSplitRuns (CJK fallback)", () => {
  it("is a no-op when no fallback font is configured", () => {
    setActiveCjkFallback(null);
    const runs: Run[] = [{ text: "abc世界xyz", style: CHAR }];
    expect(scriptSplitRuns(runs)).toBe(runs);
  });

  it("splits a mixed run into Latin + CJK pieces, retargeting CJK to the fallback", () => {
    setActiveCjkFallback("NotoCJK");
    const runs: Run[] = [{ text: "abc世界xyz", style: CHAR }];
    const split = scriptSplitRuns(runs);
    expect(split.map((r) => r.text)).toEqual(["abc", "世界", "xyz"]);
    expect(split[0]!.style.fontFamily).toBe("Arial, sans-serif");
    expect(split[1]!.style.fontFamily).toBe("NotoCJK"); // CJK piece retargeted
    expect(split[2]!.style.fontFamily).toBe("Arial, sans-serif");
    // Offset-transparent: concatenation + length preserved.
    expect(split.map((r) => r.text).join("")).toBe("abc世界xyz");
  });

  it("leaves a pure-Latin run untouched even with a fallback configured", () => {
    setActiveCjkFallback("NotoCJK");
    const runs: Run[] = [{ text: "hello", style: CHAR }];
    expect(scriptSplitRuns(runs)).toBe(runs);
  });
});

describe("CJK line breaking through the engine", () => {
  it("breaks long CJK text between characters (no spaces needed)", () => {
    // 40 CJK chars at 8px = 320px; content box 624px → wraps once or more.
    const text = "今".repeat(120); // 今×120
    const p = para([{ text, style: CHAR }]);
    const lines = createLayoutEngine().layout(doc(p)).pages[0]!.blocks[0]!.lines;
    expect(lines.length).toBeGreaterThan(1);
    // Every line carries text (broke mid-run, not one giant overflowing line).
    expect(lines.every((l) => l.fragments.length > 0)).toBe(true);
  });
});
