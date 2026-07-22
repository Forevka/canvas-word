// buildFontPayload turns the flat host specs (family + sizing + per-style bytes) the
// ClearScript C# side marshals into the CustomFontPayload the export pipeline consumes.
// These lock the payload shape the host depends on: one def per family (with only the
// present styles' face markers), face bytes keyed by family+style, and a hard throw on
// an unregistrable spec (so a bad font never silently substitutes to a clone).

import { describe, expect, it } from "vitest";
import { buildFontPayload, type CustomFontSpec } from "./fontPayload";

const bytes = (n: number): Uint8Array => new Uint8Array(n).fill(1);

const regularOnly: CustomFontSpec = {
  family: "Brand Sans",
  ascent: 0.9,
  descent: 0.25,
  regular: bytes(8),
};

describe("buildFontPayload", () => {
  it("builds one def + one face from a regular-only spec", () => {
    const payload = buildFontPayload([regularOnly]);
    expect(payload.defs.disableBuiltin).toEqual([]);
    expect(payload.defs.fonts).toHaveLength(1);
    const def = payload.defs.fonts[0]!;
    expect(def.family).toBe("Brand Sans");
    expect(def.sizing).toEqual({ ascent: 0.9, descent: 0.25 });
    // Only the regular marker is present (no bold/italic/boldItalic keys).
    expect(Object.keys(def.faces)).toEqual(["regular"]);
    expect(payload.faces).toEqual([{ family: "Brand Sans", style: "Regular", bytes: regularOnly.regular }]);
  });

  it("adds a face + marker for each supplied optional style", () => {
    const b = bytes(4);
    const i = bytes(5);
    const payload = buildFontPayload([{ ...regularOnly, bold: b, italic: i, label: "Brand" }]);
    const def = payload.defs.fonts[0]!;
    expect(def.label).toBe("Brand");
    expect(def.faces.bold).toBeDefined();
    expect(def.faces.italic).toBeDefined();
    expect(def.faces.boldItalic).toBeUndefined();
    expect(payload.faces.map((f) => f.style)).toEqual(["Regular", "Bold", "Italic"]);
    expect(payload.faces.find((f) => f.style === "Bold")?.bytes).toBe(b);
  });

  it("handles multiple fonts independently", () => {
    const payload = buildFontPayload([regularOnly, { ...regularOnly, family: "Brand Serif" }]);
    expect(payload.defs.fonts.map((d) => d.family)).toEqual(["Brand Sans", "Brand Serif"]);
    expect(payload.faces.map((f) => f.family)).toEqual(["Brand Sans", "Brand Serif"]);
  });

  it("throws on invalid sizing (ascent must be > 0)", () => {
    expect(() => buildFontPayload([{ ...regularOnly, ascent: 0 }])).toThrow(/not registrable/);
  });

  it("throws on a missing regular face", () => {
    const bad = { family: "X", ascent: 0.9, descent: 0.25 } as unknown as CustomFontSpec;
    expect(() => buildFontPayload([bad])).toThrow(/regular face bytes/);
  });
});
