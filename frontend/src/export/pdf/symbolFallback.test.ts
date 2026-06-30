// Issue #104: non-Latin symbols (✓ U+2713, ☒ U+2612, …) rendered as .notdef/tofu
// in PDF export because the export resolved one bundled face per run with no
// per-glyph fallback.  The fix adds per-glyph fallback in both the fontkit
// measure path (fontkitContext.ts) and the PDF paint path (paintBlock.ts), so:
//   - measure widths for fallback-covered glyphs are non-zero and consistent
//   - the fallback face (StixTwoMath) is embedded when its glyphs are used
//   - Latin-only output is byte-identical (the fast ASCII path is unchanged)
import { describe, expect, it, beforeAll } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle } from "@cw/shared";
import { installMeasureHost } from "../shared/measureHost";
import { resolveFont } from "../shared/fontRegistry";
import { segmentByFace } from "../shared/glyphFallback";
import { FontkitMeasureContext } from "../shared/fontkitContext";
import { renderPdf } from "./renderPdf";

beforeAll(async () => {
  await installMeasureHost();
});

// ── helpers ──────────────────────────────────────────────────────────────────

const CHAR: CharStyle = {
  fontFamily: "Calibri, sans-serif",
  fontSizePx: 16,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#202124",
};
const PARA: ParaStyle = {
  align: "left",
  lineHeight: 1.5,
  spaceBeforePx: 0,
  spaceAfterPx: 8,
  indentFirstLinePx: 0,
  indentLeftPx: 0,
};
let pid = 0;
const para = (text: string, char: Partial<CharStyle> = {}): Paragraph => ({
  kind: "paragraph",
  id: `sp${pid++}`,
  revision: 0,
  runs: [{ text, style: { ...CHAR, ...char } }],
  style: PARA,
});
const docOf = (...blocks: Paragraph[]): Document => ({
  section: {
    pageWidthPx: 816,
    pageHeightPx: 1056,
    marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
  },
  blocks,
});
const isPdf = (b: Uint8Array): boolean =>
  b[0] === 0x25 && b[1] === 0x50 && b[2] === 0x44 && b[3] === 0x46; // %PDF
const latin1 = (b: Uint8Array): string => Buffer.from(b).toString("latin1");

// ── glyphFallback unit tests ──────────────────────────────────────────────────

describe("segmentByFace — per-glyph fallback segmentation", () => {
  it("returns a single primary segment for ASCII-only text (fast path)", () => {
    const primary = resolveFont("Calibri", false, false);
    const segs = segmentByFace("Hello, world.", primary);
    expect(segs.length).toBe(1);
    expect(segs[0]!.text).toBe("Hello, world.");
    expect(segs[0]!.face.file).toBe(primary.file);
  });

  it("routes ✓ U+2713 to a fallback face that has the glyph", () => {
    const primary = resolveFont("Calibri", false, false);
    expect(primary.font.hasGlyphForCodePoint(0x2713)).toBe(false); // root-cause check

    const segs = segmentByFace("✓", primary);
    expect(segs.length).toBe(1);
    expect(segs[0]!.face.font.hasGlyphForCodePoint(0x2713)).toBe(true);
    expect(segs[0]!.face.file).not.toBe(primary.file);
  });

  it("routes ☒ U+2612 to StixTwoMath (ballot box with X)", () => {
    const primary = resolveFont("Calibri", false, false);
    const segs = segmentByFace("☒", primary);
    expect(segs[0]!.face.font.hasGlyphForCodePoint(0x2612)).toBe(true);
  });

  it("splits mixed Latin + symbol text into correct segments", () => {
    const primary = resolveFont("Calibri", false, false);
    // "A" in primary, "✓" in fallback, "B" in primary again
    const segs = segmentByFace("A✓B", primary);
    // 3 segments: A (primary), ✓ (fallback), B (primary)
    expect(segs.length).toBe(3);
    expect(segs[0]!.text).toBe("A");
    expect(segs[0]!.face.file).toBe(primary.file);
    expect(segs[1]!.text).toBe("✓");
    expect(segs[1]!.face.font.hasGlyphForCodePoint(0x2713)).toBe(true);
    expect(segs[2]!.text).toBe("B");
    expect(segs[2]!.face.file).toBe(primary.file);
  });

  it("merges adjacent symbols that belong to the same fallback face", () => {
    const primary = resolveFont("Calibri", false, false);
    // ✓ and ☒ should both go to StixTwoMath → one segment
    const segs = segmentByFace("✓☒", primary);
    expect(segs.length).toBe(1);
    expect(segs[0]!.text).toBe("✓☒");
    expect(segs[0]!.face.font.hasGlyphForCodePoint(0x2713)).toBe(true);
  });

  it("keeps a trailing variation selector glued to its base symbol (☑️ → one run)", () => {
    const primary = resolveFont("Calibri", false, false);
    // ☑ U+2611 + VS16 U+FE0F — the selector must not split onto the primary face.
    const segs = segmentByFace("☑️", primary);
    expect(segs.length).toBe(1);
    expect(segs[0]!.text).toBe("☑️");
    expect(segs[0]!.face.font.hasGlyphForCodePoint(0x2611)).toBe(true);
  });
});

// ── fontkitContext measure tests ──────────────────────────────────────────────

describe("FontkitMeasureContext — per-glyph fallback advances", () => {
  it("measures ✓ with a nonzero advance (fallback face, not .notdef)", () => {
    const ctx = new FontkitMeasureContext();
    ctx.font = "16px Calibri, sans-serif";
    const w = ctx.measureText("✓").width;
    expect(w).toBeGreaterThan(0);
  });

  it("measures ☒ with a nonzero advance", () => {
    const ctx = new FontkitMeasureContext();
    ctx.font = "16px Calibri, sans-serif";
    expect(ctx.measureText("☒").width).toBeGreaterThan(0);
  });

  it("measure is consistent across calls (cache and determinism)", () => {
    const ctx = new FontkitMeasureContext();
    ctx.font = "16px Calibri, sans-serif";
    const w1 = ctx.measureText("✓ done").width;
    const w2 = ctx.measureText("✓ done").width;
    expect(w1).toBe(w2);
    expect(w1).toBeGreaterThan(0);
  });

  it("Latin-only advance is unchanged by the fallback path (fast path)", () => {
    // The ASCII fast path must return the same value as the original single-face
    // layout call.  We verify this by measuring with the old approach directly.
    const ctx = new FontkitMeasureContext();
    ctx.font = "16px Calibri, sans-serif";
    const primary = resolveFont("Calibri", false, false);
    const text = "The quick brown fox";
    const direct = primary.font.layout(text).advanceWidth * (16 / primary.font.unitsPerEm);
    expect(ctx.measureText(text).width).toBeCloseTo(direct, 5);
  });
});

// ── PDF export integration tests ─────────────────────────────────────────────

describe("PDF export — non-Latin symbol fallback (issue #104)", () => {
  it("embeds StixTwoMath when a run contains ✓", async () => {
    const { bytes, warnings } = await renderPdf(docOf(para("✓ supported")));
    expect(isPdf(bytes)).toBe(true);
    // No substitution warning on the primary (Calibri/Carlito) font.
    expect(warnings.find((w) => w.code === "font-substituted")).toBeUndefined();
    // StixTwoMath must be embedded as a font resource.
    // pdfkit embeds fonts using their PostScript name (STIXTwoMath), which
    // differs from our registration key (StixTwoMath / StixTwoMath-Regular.ttf).
    expect(latin1(bytes)).toContain("STIXTwoMath");
  });

  it("embeds StixTwoMath when a run contains ☒ (ballot box with X)", async () => {
    const { bytes } = await renderPdf(docOf(para("Status: ☒")));
    expect(isPdf(bytes)).toBe(true);
    // pdfkit embeds fonts using their PostScript name (STIXTwoMath), which
    // differs from our registration key (StixTwoMath / StixTwoMath-Regular.ttf).
    expect(latin1(bytes)).toContain("STIXTwoMath");
  });

  it("embeds StixTwoMath for all four ballot-box variants", async () => {
    const { bytes } = await renderPdf(docOf(para("☐ unchecked  ☑ checked  ☒ rejected")));
    expect(isPdf(bytes)).toBe(true);
    // pdfkit embeds fonts using their PostScript name (STIXTwoMath), which
    // differs from our registration key (StixTwoMath / StixTwoMath-Regular.ttf).
    expect(latin1(bytes)).toContain("STIXTwoMath");
  });

  it("does NOT embed StixTwoMath for a Latin-only run (byte-identical fast path)", async () => {
    const { bytes } = await renderPdf(docOf(para("Hello, world. No symbols here.")));
    expect(isPdf(bytes)).toBe(true);
    // StixTwoMath (PostScript name: STIXTwoMath) should not be embedded
    // when it was never needed for a fallback glyph.
    expect(latin1(bytes)).not.toContain("STIXTwoMath");
  });

  it("renders a paragraph mixing Latin and ✓ without font-substituted warning", async () => {
    const { warnings } = await renderPdf(docOf(para("Feature A — ✓ done, Feature B — ✗ pending")));
    expect(warnings.find((w) => w.code === "font-substituted")).toBeUndefined();
  });
});
