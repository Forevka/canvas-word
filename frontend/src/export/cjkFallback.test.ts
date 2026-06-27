// Bundled CJK fallback (subset of Noto Sans SC). Before this, CJK text rendered as
// .notdef/tofu/"x" in PDF export because only Latin metric-clones were bundled and
// nothing supplied a CJK face. These assert: the face is a single-Regular built-in
// with real CJK glyphs, the fallback defaults on (zero config), PDF export uses it
// (no substitution), and DOCX keeps the document's own family (no bundled-clone leak).
import { describe, expect, it, beforeAll } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import type { CharStyle, Document, Paragraph, ParaStyle } from "@cw/shared";
import { installMeasureHost } from "./shared/measureHost";
import { resolveFont } from "./shared/fontRegistry";
import { runExport } from "./pipeline";
import { CJK_FONT_FAMILY } from "../fonts/clones";
import { resolveCjk } from "../config";

beforeAll(async () => {
  await installMeasureHost();
});

// Common hanzi (all within the GB2312 Level-1 subset we ship) + CJK punctuation.
const CJK = "中文测试，你好世界。";
const ZHONG = 0x4e2d; // 中

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
  align: "left",
  lineHeight: 1.5,
  spaceBeforePx: 0,
  spaceAfterPx: 8,
  indentFirstLinePx: 0,
  indentLeftPx: 0,
};
const docOf = (text: string, char: Partial<CharStyle> = {}): Document => ({
  section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
  blocks: [
    { kind: "paragraph", id: "p0", revision: 0, runs: [{ text, style: { ...CHAR, ...char } }], style: PARA } as Paragraph,
  ],
});
const isPdf = (b: Uint8Array): boolean => b[0] === 0x25 && b[1] === 0x50 && b[2] === 0x44 && b[3] === 0x46; // %PDF

describe("bundled CJK fallback face", () => {
  it("maps every NotoSansSC style to the single Regular face, which has CJK glyphs", () => {
    for (const [bold, italic] of [
      [false, false],
      [true, false],
      [false, true],
      [true, true],
    ] as const) {
      const r = resolveFont(CJK_FONT_FAMILY, bold, italic);
      expect(r.file).toBe("NotoSansSC-Regular.ttf");
      expect(r.substituted).toBe(false);
      expect(r.font.hasGlyphForCodePoint(ZHONG)).toBe(true);
    }
  });

  it("Latin clones lack CJK glyphs (the root cause of the tofu)", () => {
    expect(resolveFont("Arial", false, false).font.hasGlyphForCodePoint(ZHONG)).toBe(false);
    expect(resolveFont("Times New Roman", false, false).font.hasGlyphForCodePoint(ZHONG)).toBe(false);
  });
});

describe("CJK fallback defaulting", () => {
  it("resolveCjk defaults the fallback to the bundled family; explicit '' opts out", () => {
    expect(resolveCjk().fallbackFont).toBe(CJK_FONT_FAMILY);
    expect(resolveCjk({}).fallbackFont).toBe(CJK_FONT_FAMILY);
    expect(resolveCjk({ fallbackFont: "MyCustomCJK" }).fallbackFont).toBe("MyCustomCJK");
    expect(resolveCjk({ fallbackFont: "" }).fallbackFont).toBe("");
  });
});

describe("export with CJK text", () => {
  it("renders a CJK paragraph to a valid PDF with no font substitution (zero config)", async () => {
    const { bytes, warnings } = await runExport(docOf(CJK), "pdf");
    expect(isPdf(bytes)).toBe(true);
    expect(bytes.length).toBeGreaterThan(500);
    expect(warnings.find((w) => w.code === "font-substituted")).toBeUndefined();
  });

  it("keeps concurrent exports from cross-contaminating CJK state", async () => {
    // The CJK fallback is threaded into each job's layout engine and re-asserted
    // synchronously per (await-free) layout pass, so overlapping jobs don't share
    // state: a default-fallback render and an opt-out render each get their own
    // setting (one embeds NotoSansSC, the other doesn't) even when interleaved.
    const decode = (b: Uint8Array): string => new TextDecoder("latin1").decode(b);
    const [withFallback, optedOut] = await Promise.all([
      runExport(docOf(CJK), "pdf"),
      runExport(docOf(CJK), "pdf", {}, undefined, { fallbackFont: "" }),
    ]);
    expect(decode(withFallback.bytes).includes(CJK_FONT_FAMILY)).toBe(true);
    expect(decode(optedOut.bytes).includes(CJK_FONT_FAMILY)).toBe(false);
  });

  it("keeps the document's own family in DOCX (CJK survives via Word; no bundled-clone leak)", async () => {
    const { bytes } = await runExport(docOf(CJK), "docx");
    const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!);
    expect(xml).toContain(CJK); // the CJK text is written verbatim
    expect(xml).toContain("Calibri"); // original family preserved on the run
    expect(xml).not.toContain(CJK_FONT_FAMILY); // the export's fallback never leaks into the .docx
  });
});
