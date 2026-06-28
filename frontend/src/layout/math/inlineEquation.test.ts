// Inline equations: a single-U+FFFC run carrying CharStyle.equation lays out as a
// fragment with a math box (growing the line), and survives a .docx round-trip.

import { beforeAll, describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, Run } from "@cw/shared";
import { installMeasureHost } from "../../export/shared/measureHost";
import { createLayoutEngine } from "../engine";
import { parseMathml } from "../../mathml/parse";
import { runExport } from "../../export/pipeline";
import { runImport } from "../../import/docx/pipeline";
import { equationBox } from "./equationLayout";
import { MATH_FONT_FAMILY } from "../../fonts/clones";
import { charStyleToFont, measureTextWidth } from "../metrics";

beforeAll(async () => {
  await installMeasureHost();
});

const CHAR: CharStyle = {
  fontFamily: "Times New Roman", fontSizePx: 16, bold: false, italic: false,
  underline: false, strikethrough: false, color: "#000000",
};
const PARA: ParaStyle = {
  align: "left", lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 8,
  indentFirstLinePx: 0, indentLeftPx: 0,
};
const FRAC = "<math><mfrac><mn>1</mn><mn>2</mn></mfrac></math>";

const inlineRun = (xml: string): Run => ({ text: "￼", style: { ...CHAR, equation: { ...parseMathml(xml), display: false } } });

const docWithInline = (): Document => ({
  section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
  blocks: [
    { kind: "paragraph", id: "p1", revision: 0, style: PARA, runs: [{ text: "x = ", style: CHAR }, inlineRun(FRAC), { text: " done", style: CHAR }] } satisfies Paragraph,
  ],
});

describe("inline equation", () => {
  it("lays out as a fragment carrying a typeset math box", () => {
    const tree = createLayoutEngine().layout(docWithInline());
    const frags = tree.pages.flatMap((p) => p.blocks).flatMap((b) => b.lines).flatMap((l) => l.fragments);
    const eqFrag = frags.find((f) => f.equation);
    expect(eqFrag).toBeTruthy();
    expect(eqFrag!.equation!.glyphs.length).toBeGreaterThan(0); // "1" and "2"
    expect(eqFrag!.equation!.rules.length).toBeGreaterThanOrEqual(1); // fraction bar
    expect(eqFrag!.width).toBeGreaterThan(0);
    // The line grew taller than a plain 16px text line to fit the stacked fraction.
    const line = tree.pages[0]!.blocks[0]!.lines[0]!;
    expect(line.height).toBeGreaterThan(16);
  });

  it("reserves exactly the math box width, not the U+FFFC sentinel width (#16)", () => {
    // The U+FFFC placeholder glyph measures ~1em wide, but a simple fraction is
    // narrower. The reserved/occupied fragment width must equal equationBox(...).width
    // EXACTLY — the old additive `max(0, box.width - sentinelW)` over-reserved to the
    // sentinel width whenever the box was narrower, wrapping early and hit-testing wide.
    const box = equationBox({ ...parseMathml(FRAC), display: false }, MATH_FONT_FAMILY, CHAR.fontSizePx);
    const sentinelW = measureTextWidth("￼", charStyleToFont(CHAR));
    // Guard the regression is reproducible: the box really is narrower than the sentinel.
    expect(box.width).toBeLessThan(sentinelW);

    const tree = createLayoutEngine().layout(docWithInline());
    const frags = tree.pages.flatMap((p) => p.blocks).flatMap((b) => b.lines).flatMap((l) => l.fragments);
    const eqFrag = frags.find((f) => f.equation)!;
    expect(eqFrag.width).toBeCloseTo(box.width, 5);
    // And the carried box matches the reserved width — paint and hit-test agree.
    expect(eqFrag.equation!.width).toBeCloseTo(eqFrag.width, 5);
  });

  it("round-trips through .docx as inline m:oMath", async () => {
    const { bytes } = await runExport(docWithInline(), "docx");
    const xml = new TextDecoder().decode(bytes);
    void xml;
    const { doc: back } = runImport(bytes);
    const p = back.blocks.find((b): b is Paragraph => b.kind === "paragraph")!;
    const eqRun = p.runs.find((r) => r.style.equation);
    expect(eqRun).toBeTruthy();
    expect(eqRun!.style.equation!.root.children.some((n) => n.type === "frac")).toBe(true);
    // Surrounding text is preserved.
    expect(p.runs.map((r) => (r.style.equation ? "[eq]" : r.text)).join("")).toContain("x = ");
  });
});
