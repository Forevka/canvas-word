import { describe, expect, it } from "vitest";
import { normalizeFontsConfig } from "./fontsConfig";

const validFont = () => ({
  family: "Acme",
  faces: { regular: "https://x/acme.woff2" },
  sizing: { ascent: 0.9, descent: 0.2 },
});

describe("normalizeFontsConfig", () => {
  it("returns undefined for non-objects", () => {
    expect(normalizeFontsConfig(null)).toBeUndefined();
    expect(normalizeFontsConfig(undefined)).toBeUndefined();
    expect(normalizeFontsConfig("nope")).toBeUndefined();
    expect(normalizeFontsConfig(42)).toBeUndefined();
    expect(normalizeFontsConfig([])).toBeUndefined();
  });

  it("returns empty arrays for an object with no fonts/disableBuiltin", () => {
    expect(normalizeFontsConfig({})).toEqual({ disableBuiltin: [], fonts: [] });
  });

  it("keeps a fully valid font with only the regular face", () => {
    const out = normalizeFontsConfig({ fonts: [validFont()] });
    expect(out?.fonts).toEqual([
      { family: "Acme", faces: { regular: "https://x/acme.woff2" }, sizing: { ascent: 0.9, descent: 0.2 } },
    ]);
  });

  it("drops a font missing family", () => {
    const f = validFont() as Record<string, unknown>;
    delete f.family;
    expect(normalizeFontsConfig({ fonts: [f] })?.fonts).toEqual([]);
  });

  it("drops a font with empty-string family", () => {
    expect(normalizeFontsConfig({ fonts: [{ ...validFont(), family: "" }] })?.fonts).toEqual([]);
  });

  it("drops a font missing faces.regular", () => {
    expect(normalizeFontsConfig({ fonts: [{ ...validFont(), faces: {} }] })?.fonts).toEqual([]);
  });

  it("drops a font whose faces is not an object", () => {
    expect(normalizeFontsConfig({ fonts: [{ ...validFont(), faces: "x" }] })?.fonts).toEqual([]);
  });

  it("drops a font missing sizing", () => {
    const f = validFont() as Record<string, unknown>;
    delete f.sizing;
    expect(normalizeFontsConfig({ fonts: [f] })?.fonts).toEqual([]);
  });

  it("drops a font with ascent <= 0", () => {
    expect(normalizeFontsConfig({ fonts: [{ ...validFont(), sizing: { ascent: 0, descent: 0.2 } }] })?.fonts).toEqual([]);
  });

  it("drops a font with negative descent", () => {
    expect(normalizeFontsConfig({ fonts: [{ ...validFont(), sizing: { ascent: 0.9, descent: -0.1 } }] })?.fonts).toEqual([]);
  });

  it("accepts descent of exactly 0", () => {
    const out = normalizeFontsConfig({ fonts: [{ ...validFont(), sizing: { ascent: 0.9, descent: 0 } }] });
    expect(out?.fonts).toHaveLength(1);
    expect(out?.fonts[0]!.sizing).toEqual({ ascent: 0.9, descent: 0 });
  });

  it("drops a font with non-finite sizing values", () => {
    expect(
      normalizeFontsConfig({ fonts: [{ ...validFont(), sizing: { ascent: Number.POSITIVE_INFINITY, descent: 0 } }] })
        ?.fonts,
    ).toEqual([]);
    expect(
      normalizeFontsConfig({ fonts: [{ ...validFont(), sizing: { ascent: 0.9, descent: Number.NaN } }] })?.fonts,
    ).toEqual([]);
  });

  it("drops a font whose sizing values are non-numbers", () => {
    expect(
      normalizeFontsConfig({ fonts: [{ ...validFont(), sizing: { ascent: "0.9", descent: 0 } }] })?.fonts,
    ).toEqual([]);
  });

  it("keeps optional bold/italic/boldItalic faces and label when present", () => {
    const out = normalizeFontsConfig({
      fonts: [
        {
          ...validFont(),
          faces: {
            regular: "r.woff2",
            bold: "b.woff2",
            italic: "i.woff2",
            boldItalic: "bi.woff2",
          },
          label: "Acme Display",
        },
      ],
    });
    expect(out?.fonts[0]!.faces).toEqual({
      regular: "r.woff2",
      bold: "b.woff2",
      italic: "i.woff2",
      boldItalic: "bi.woff2",
    });
    expect(out?.fonts[0]!.label).toBe("Acme Display");
  });

  it("drops empty-string optional faces and label rather than keeping them", () => {
    const out = normalizeFontsConfig({
      fonts: [{ ...validFont(), faces: { regular: "r.woff2", bold: "" }, label: "" }],
    });
    expect(out?.fonts[0]!.faces).toEqual({ regular: "r.woff2" });
    expect("label" in out!.fonts[0]!).toBe(false);
  });

  it("filters non-string entries out of disableBuiltin", () => {
    const out = normalizeFontsConfig({ disableBuiltin: ["Arial", 5, null, "Calibri", {}] });
    expect(out?.disableBuiltin).toEqual(["Arial", "Calibri"]);
  });

  it("defaults disableBuiltin to [] when it is not an array", () => {
    expect(normalizeFontsConfig({ disableBuiltin: "Arial" })?.disableBuiltin).toEqual([]);
  });

  it("keeps only valid fonts when a mix is supplied", () => {
    const out = normalizeFontsConfig({
      fonts: [validFont(), { family: "Bad" }, { ...validFont(), family: "Good2" }],
    });
    expect(out?.fonts.map((f) => f.family)).toEqual(["Acme", "Good2"]);
  });
});
