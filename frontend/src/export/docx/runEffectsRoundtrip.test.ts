// Minor run typography & effects (issue #60) — every w:rPr extra must survive a
// full .docx round-trip (export → re-import): w:dstrike, w:position, w:kern, w:w,
// w:em, w:outline/w:shadow/w:emboss/w:imprint, w:bdr, and w:fitText.
import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, Run, SectionProps } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import { writeDocx } from "./writeDocx";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#202124" };
const PARA = { align: "left" as const, lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

const styledRun = (text: string, patch: Partial<CharStyle>): Run => ({ text, style: { ...CHAR, ...patch } });

const roundTrip = (runs: Run[]): Paragraph => {
  const doc: Document = { section: SECTION, blocks: [{ kind: "paragraph", id: "p1", revision: 0, style: { ...PARA }, runs }] };
  return runImport(writeDocx(doc).bytes).doc.blocks[0] as Paragraph;
};

describe("minor run effects .docx round-trip", () => {
  it("preserves double strikethrough (w:dstrike)", () => {
    const p = roundTrip([styledRun("dstrike", { doubleStrikethrough: true }), styledRun("plain", {})]);
    expect(p.runs.find((r) => r.text === "dstrike")!.style.doubleStrikethrough).toBe(true);
    expect(p.runs.find((r) => r.text === "plain")!.style.doubleStrikethrough).toBeUndefined();
  });

  it("preserves a raised AND a lowered baseline position (w:position, signed)", () => {
    const p = roundTrip([styledRun("up", { positionPx: 4 }), styledRun("down", { positionPx: -4 })]);
    expect(p.runs.find((r) => r.text === "up")!.style.positionPx).toBe(4);
    expect(p.runs.find((r) => r.text === "down")!.style.positionPx).toBe(-4);
  });

  it("preserves the kerning threshold (w:kern)", () => {
    const p = roundTrip([styledRun("kern", { kerningMinPx: 12 })]);
    expect(p.runs[0]!.style.kerningMinPx).toBe(12);
  });

  it("preserves character width scaling (w:w)", () => {
    const p = roundTrip([styledRun("wide", { widthScalePct: 150 }), styledRun("narrow", { widthScalePct: 66 })]);
    expect(p.runs.find((r) => r.text === "wide")!.style.widthScalePct).toBe(150);
    expect(p.runs.find((r) => r.text === "narrow")!.style.widthScalePct).toBe(66);
  });

  it("preserves emphasis marks (w:em)", () => {
    const p = roundTrip([
      styledRun("dot", { emphasisMark: "dot" }),
      styledRun("circle", { emphasisMark: "circle" }),
      styledRun("underDot", { emphasisMark: "underDot" }),
    ]);
    expect(p.runs.find((r) => r.text === "dot")!.style.emphasisMark).toBe("dot");
    expect(p.runs.find((r) => r.text === "circle")!.style.emphasisMark).toBe("circle");
    expect(p.runs.find((r) => r.text === "underDot")!.style.emphasisMark).toBe("underDot");
  });

  it("preserves the outline/shadow/emboss/imprint text effects", () => {
    const p = roundTrip([
      styledRun("outline", { outline: true }),
      styledRun("shadow", { shadow: true }),
      styledRun("emboss", { emboss: true }),
      styledRun("imprint", { imprint: true }),
    ]);
    expect(p.runs.find((r) => r.text === "outline")!.style.outline).toBe(true);
    expect(p.runs.find((r) => r.text === "shadow")!.style.shadow).toBe(true);
    expect(p.runs.find((r) => r.text === "emboss")!.style.emboss).toBe(true);
    expect(p.runs.find((r) => r.text === "imprint")!.style.imprint).toBe(true);
  });

  it("preserves an explicit snap-to-grid OFF (w:snapToGrid, issue #161)", () => {
    // Word's default is ON, so the meaningful override is the explicit `false`
    // (w:val="0"); a plain run must not gain the field.
    const p = roundTrip([styledRun("gridOff", { snapToGrid: false }), styledRun("plain", {})]);
    expect(p.runs.find((r) => r.text === "gridOff")!.style.snapToGrid).toBe(false);
    expect(p.runs.find((r) => r.text === "plain")!.style.snapToGrid).toBeUndefined();
  });

  it("preserves a run border (w:bdr), including its line style", () => {
    const p = roundTrip([styledRun("bordered", { runBorder: { color: "#1a73e8", widthPx: 1.5, style: "dashed" } })]);
    const b = p.runs[0]!.style.runBorder;
    expect(b).toEqual({ color: "#1a73e8", widthPx: 1.5, style: "dashed" });
  });

  it("preserves a fitText target width (w:fitText)", () => {
    const p = roundTrip([styledRun("fit", { fitTextPx: 60 })]);
    expect(p.runs[0]!.style.fitTextPx).toBe(60);
  });

  it("keeps a plain run untouched (no effect fields leak in)", () => {
    const p = roundTrip([styledRun("plain", {})]);
    const s = p.runs[0]!.style;
    expect(s.doubleStrikethrough).toBeUndefined();
    expect(s.positionPx).toBeUndefined();
    expect(s.widthScalePct).toBeUndefined();
    expect(s.emphasisMark).toBeUndefined();
    expect(s.snapToGrid).toBeUndefined();
    expect(s.runBorder).toBeUndefined();
    expect(s.fitTextPx).toBeUndefined();
  });
});
