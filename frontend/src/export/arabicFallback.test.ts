// Bundled Arabic fallback (Noto Sans Arabic). Before this, Arabic text rendered as
// .notdef/tofu/"x" in PDF export because only Latin metric-clones were bundled and
// nothing supplied an Arabic face with contextual-shaping support. These assert:
// the face is a single-Regular built-in with real Arabic glyphs + shaping, the
// fallback defaults on (zero config), PDF export uses it (no substitution), DOCX
// keeps the document's own family (no bundled-clone leak), and Arabic text measures
// non-zero with the correct face.
import { describe, expect, it, beforeAll } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import type { CharStyle, Document, Paragraph, ParaStyle } from "@cw/shared";
import { installMeasureHost } from "./shared/measureHost";
import { resolveFont } from "./shared/fontRegistry";
import { runExport } from "./pipeline";
import { ARABIC_FONT_FAMILY } from "../fonts/clones";

beforeAll(async () => {
  await installMeasureHost();
});

// A short Arabic sentence with varied letter positions (initial/medial/final/isolated
// forms) so shaping is exercised. "مرحبا بالعالم" = "Hello World".
const ARABIC = "مرحبا بالعالم";
const MIM = 0x0645; // م — a Basic Arabic letter

const CHAR: CharStyle = {
  fontFamily: "Calibri, sans-serif",
  fontSizePx: 16,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#000000",
};
const PARA: ParaStyle = {
  align: "right",
  lineHeight: 1.5,
  spaceBeforePx: 0,
  spaceAfterPx: 8,
  indentFirstLinePx: 0,
  indentLeftPx: 0,
  direction: "rtl",
};
const docOf = (text: string, char: Partial<CharStyle> = {}): Document => ({
  section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
  blocks: [
    { kind: "paragraph", id: "p0", revision: 0, runs: [{ text, style: { ...CHAR, ...char } }], style: PARA } as Paragraph,
  ],
});
const isPdf = (b: Uint8Array): boolean => b[0] === 0x25 && b[1] === 0x50 && b[2] === 0x44 && b[3] === 0x46;

describe("bundled Arabic fallback face", () => {
  it("maps every NotoSansArabic style to the single Regular face, which has Arabic glyphs", () => {
    for (const [bold, italic] of [
      [false, false],
      [true, false],
      [false, true],
      [true, true],
    ] as const) {
      const r = resolveFont(ARABIC_FONT_FAMILY, bold, italic);
      expect(r.file).toBe("NotoSansArabic-Regular.ttf");
      expect(r.substituted).toBe(false);
      expect(r.font.hasGlyphForCodePoint(MIM)).toBe(true);
    }
  });

  it("Arabic face shapes the run to non-zero advance with at least one glyph", () => {
    const { font } = resolveFont(ARABIC_FONT_FAMILY, false, false);
    // fontkit.layout() applies GSUB (joining forms); the advance width must be > 0
    // for the test string (i.e. the shaped run doesn't collapse to zero-width tofu).
    const run = font.layout(ARABIC);
    expect(run.advanceWidth).toBeGreaterThan(0);
    expect(run.glyphs.length).toBeGreaterThan(0);
  });

  it("Latin clones lack Arabic glyphs (the root cause of the tofu)", () => {
    expect(resolveFont("Arial", false, false).font.hasGlyphForCodePoint(MIM)).toBe(false);
    expect(resolveFont("Calibri", false, false).font.hasGlyphForCodePoint(MIM)).toBe(false);
  });
});

describe("export with Arabic text", () => {
  it("renders an Arabic paragraph to a valid PDF with no font substitution (zero config)", async () => {
    const { bytes, warnings } = await runExport(docOf(ARABIC), "pdf");
    expect(isPdf(bytes)).toBe(true);
    expect(bytes.length).toBeGreaterThan(500);
    expect(warnings.find((w) => w.code === "font-substituted")).toBeUndefined();
  });

  it("embeds the Arabic face in the PDF (family name present in the byte stream)", async () => {
    const { bytes } = await runExport(docOf(ARABIC), "pdf");
    const text = new TextDecoder("latin1").decode(bytes);
    expect(text.includes(ARABIC_FONT_FAMILY)).toBe(true);
  });

  it("keeps the document's own family in DOCX (Arabic survives via renderer; no bundled-clone leak)", async () => {
    const { bytes } = await runExport(docOf(ARABIC), "docx");
    const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
    expect(xml).toContain(ARABIC);
    expect(xml).toContain("Calibri");
    expect(xml).not.toContain(ARABIC_FONT_FAMILY);
  });

  it("opt-out (arabicFallbackFont:'') does NOT embed the Arabic face", async () => {
    const { bytes } = await runExport(docOf(ARABIC), "pdf", {}, undefined, { arabicFallbackFont: "" });
    const text = new TextDecoder("latin1").decode(bytes);
    expect(text.includes(ARABIC_FONT_FAMILY)).toBe(false);
  });
});
