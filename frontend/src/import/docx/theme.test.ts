// theme1.xml parsing: font scheme (major/minor latin typefaces), color scheme
// (slot -> hex), plus the themeColor alias map and themeFont slot resolution.

import { describe, expect, it } from "vitest";
import { EMPTY_THEME, parseThemeXml, themeColor, themeFont } from "./theme";

const THEME_XML = `<?xml version="1.0"?>
<a:theme xmlns:a="urn">
  <a:themeElements>
    <a:fontScheme>
      <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface="Yu Gothic Light"/><a:cs typeface="Times New Roman"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="Yu Gothic"/><a:cs typeface="Arial"/></a:minorFont>
    </a:fontScheme>
    <a:clrScheme>
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>`;

describe("parseThemeXml", () => {
  it("extracts major and minor latin typefaces", () => {
    const t = parseThemeXml(THEME_XML);
    expect(t.majorLatin).toBe("Calibri Light");
    expect(t.minorLatin).toBe("Calibri");
  });

  it("extracts the East-Asian (a:ea) and complex-script (a:cs) typefaces", () => {
    const t = parseThemeXml(THEME_XML);
    expect(t.majorEastAsia).toBe("Yu Gothic Light");
    expect(t.minorEastAsia).toBe("Yu Gothic");
    expect(t.majorComplexScript).toBe("Times New Roman");
    expect(t.minorComplexScript).toBe("Arial");
  });

  it("extracts srgbClr scheme colors by slot name", () => {
    const t = parseThemeXml(THEME_XML);
    expect(t.colors.get("dk2")).toBe("44546A");
    expect(t.colors.get("accent1")).toBe("4472C4");
  });

  it("falls back to sysClr lastClr when there is no srgbClr", () => {
    const t = parseThemeXml(THEME_XML);
    expect(t.colors.get("dk1")).toBe("000000");
    expect(t.colors.get("lt1")).toBe("FFFFFF");
  });

  it("returns an empty theme when themeElements is missing", () => {
    const t = parseThemeXml(`<a:theme xmlns:a="urn"></a:theme>`);
    expect(t.majorLatin).toBeUndefined();
    expect(t.minorLatin).toBeUndefined();
    expect(t.colors.size).toBe(0);
  });

  it("ignores an empty typeface=\"\"", () => {
    const t = parseThemeXml(
      `<a:theme xmlns:a="urn"><a:themeElements><a:fontScheme>` +
        `<a:majorFont><a:latin typeface=""/></a:majorFont>` +
        `</a:fontScheme></a:themeElements></a:theme>`,
    );
    expect(t.majorLatin).toBeUndefined();
  });
});

describe("themeColor", () => {
  const theme = parseThemeXml(THEME_XML);

  it("resolves an aliased name to the underlying scheme slot", () => {
    // text1 -> dk1, background1 -> lt1, hyperlink -> hlink
    expect(themeColor(theme, "text1")).toBe("000000");
    expect(themeColor(theme, "background1")).toBe("FFFFFF");
    expect(themeColor(theme, "hyperlink")).toBe("0563C1");
  });

  it("resolves a direct (non-aliased) slot name", () => {
    expect(themeColor(theme, "accent1")).toBe("4472C4");
  });

  it("returns undefined for an unknown slot", () => {
    expect(themeColor(theme, "accent6")).toBeUndefined();
  });
});

describe("themeFont", () => {
  const theme = parseThemeXml(THEME_XML);

  it("maps major HAnsi/Ascii slots to the major latin typeface", () => {
    expect(themeFont(theme, "majorHAnsi")).toBe("Calibri Light");
    expect(themeFont(theme, "majorAscii")).toBe("Calibri Light");
  });

  it("maps minor HAnsi/Ascii slots to the minor latin typeface", () => {
    expect(themeFont(theme, "minorHAnsi")).toBe("Calibri");
  });

  it("maps EastAsia / Bidi slots to their own (non-latin) theme faces", () => {
    expect(themeFont(theme, "majorEastAsia")).toBe("Yu Gothic Light");
    expect(themeFont(theme, "minorEastAsia")).toBe("Yu Gothic");
    expect(themeFont(theme, "majorBidi")).toBe("Times New Roman");
    expect(themeFont(theme, "minorBidi")).toBe("Arial");
  });

  it("returns undefined for a non major/minor slot", () => {
    expect(themeFont(theme, "somethingElse")).toBeUndefined();
  });
});

describe("EMPTY_THEME", () => {
  it("is an empty, color-less theme", () => {
    expect(EMPTY_THEME.majorLatin).toBeUndefined();
    expect(EMPTY_THEME.colors.size).toBe(0);
  });
});
