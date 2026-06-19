import { describe, it, expect } from "vitest";
import { DEFAULT_CHAR_STYLE, DEFAULT_PARA_STYLE, makeDefaultCharStyle, makeDefaultParaStyle } from "./defaults";
import { defaultStylesheet, makeDefaultStylesheet, styleById } from "./stylesheet";

describe("default typography factories", () => {
  it("zero-arg makeDefaultCharStyle/ParaStyle equal the exported constants", () => {
    expect(makeDefaultCharStyle()).toEqual(DEFAULT_CHAR_STYLE);
    expect(makeDefaultParaStyle()).toEqual(DEFAULT_PARA_STYLE);
  });

  it("applies char overrides without touching the other fields", () => {
    const c = makeDefaultCharStyle({ fontFamily: "Times New Roman, serif", fontSizePx: 14, color: "#111111" });
    expect(c.fontFamily).toBe("Times New Roman, serif");
    expect(c.fontSizePx).toBe(14);
    expect(c.color).toBe("#111111");
    // Untouched flags stay at the built-in defaults.
    expect(c.bold).toBe(false);
    expect(c.underline).toBe(false);
  });

  it("applies the paragraph line-height override", () => {
    expect(makeDefaultParaStyle({ lineHeight: 2 }).lineHeight).toBe(2);
    expect(makeDefaultParaStyle().lineHeight).toBe(1.5);
  });
});

describe("makeDefaultStylesheet", () => {
  it("zero-arg reproduces defaultStylesheet() byte-for-byte", () => {
    expect(makeDefaultStylesheet()).toEqual(defaultStylesheet());
  });

  it("overrides the Normal body font/size/color", () => {
    const sheet = makeDefaultStylesheet({ fontFamily: "Times New Roman, serif", fontSizePx: 14, color: "#101010" });
    const normal = styleById(sheet, "Normal")!;
    expect(normal.char.fontFamily).toBe("Times New Roman, serif");
    expect(normal.char.fontSizePx).toBe(14);
    expect(normal.char.color).toBe("#101010");
  });

  it("retargets the heading font family but keeps heading sizes", () => {
    const sheet = makeDefaultStylesheet({ headingFontFamily: "Calibri, sans-serif" });
    expect(styleById(sheet, "Title")!.char.fontFamily).toBe("Calibri, sans-serif");
    expect(styleById(sheet, "Heading1")!.char.fontFamily).toBe("Calibri, sans-serif");
    expect(styleById(sheet, "Heading1")!.char.fontSizePx).toBe(24);
    // Body font is independent of the heading font.
    expect(styleById(sheet, "Normal")!.char.fontFamily).toBe("Georgia, serif");
  });
});
