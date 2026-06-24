// Unit coverage for the custom-font overlay: resolution shadows built-ins, metrics
// come from the caller, the toolbar is filtered/extended per config, and invalid
// input (bad sizing, WOFF2) is rejected.

import { afterEach, describe, expect, it, vi } from "vitest";
import { cloneFamilyFor, metricsFor, toolbarFonts, CLONE_METRICS } from "./clones";
import {
  __resetCustomFonts,
  allCustomFonts,
  createFontRegistry,
  customFontFileName,
  customMetrics,
  defaultFontRegistry,
  registerCustomFonts,
  setActiveFontRegistry,
  type CustomFontDef,
} from "./customRegistry";

const inter: CustomFontDef = {
  family: "Inter",
  faces: { regular: "https://x/Inter-Regular.ttf", bold: "https://x/Inter-Bold.ttf" },
  sizing: { ascent: 0.95, descent: 0.24 },
};

afterEach(() => {
  __resetCustomFonts();
  setActiveFontRegistry(defaultFontRegistry());
  vi.restoreAllMocks();
});

describe("custom font overlay", () => {
  it("resolves a registered custom family to itself (not substituted)", () => {
    registerCustomFonts({ fonts: [inter] });
    expect(cloneFamilyFor("Inter")).toEqual({ clone: "Inter", substituted: false });
    // case/quote-insensitive, like built-in resolution
    expect(cloneFamilyFor(' "inter" ')).toEqual({ clone: "Inter", substituted: false });
  });

  it("leaves built-in resolution intact for unregistered families", () => {
    expect(cloneFamilyFor("Calibri")).toEqual({ clone: "Carlito", substituted: false });
    expect(cloneFamilyFor("Nonesuch")).toEqual({ clone: "Arimo", substituted: true });
  });

  it("returns caller-supplied metrics for a custom family, baked for clones", () => {
    registerCustomFonts({ fonts: [inter] });
    expect(customMetrics("Inter")).toEqual({ ascent: 0.95, descent: 0.24 });
    expect(metricsFor("Inter")).toEqual({ ascent: 0.95, descent: 0.24 });
    expect(metricsFor("Carlito")).toEqual(CLONE_METRICS.Carlito);
  });

  it("a custom family named like a built-in shadows the clone", () => {
    vi.spyOn(console, "warn").mockImplementation(() => {});
    registerCustomFonts({ fonts: [{ ...inter, family: "Calibri" }] });
    expect(cloneFamilyFor("Calibri")).toEqual({ clone: "Calibri", substituted: false });
    expect(metricsFor("Calibri")).toEqual({ ascent: 0.95, descent: 0.24 });
  });

  it("builds a stable synthetic file name per family+style", () => {
    expect(customFontFileName("Inter", "Bold")).toBe("__custom__inter-Bold.ttf");
    expect(customFontFileName(" Inter ", "Regular")).toBe("__custom__inter-Regular.ttf");
  });

  it("rejects invalid sizing and WOFF2, and warns on conflicting redefine", () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    registerCustomFonts({ fonts: [{ ...inter, family: "BadSize", sizing: { ascent: 0, descent: 0.2 } }] });
    registerCustomFonts({ fonts: [{ ...inter, family: "Woff2", faces: { regular: "https://x/F.woff2" } }] });
    registerCustomFonts({ fonts: [{ ...inter, family: "Woff1", faces: { regular: "https://x/F.woff" } }] });
    expect(allCustomFonts().map((f) => f.family)).not.toContain("BadSize");
    expect(allCustomFonts().map((f) => f.family)).not.toContain("Woff2"); // Brotli-wrapped SFNT
    expect(allCustomFonts().map((f) => f.family)).not.toContain("Woff1"); // zlib-wrapped SFNT

    registerCustomFonts({ fonts: [inter] });
    registerCustomFonts({ fonts: [{ ...inter, sizing: { ascent: 0.8, descent: 0.2 } }] });
    expect(customMetrics("Inter")).toEqual({ ascent: 0.8, descent: 0.2 }); // latest wins
    expect(warn).toHaveBeenCalled();
  });

  it("toolbar: built-ins minus disableBuiltin, plus custom (per-instance)", () => {
    registerCustomFonts({ fonts: [inter] });
    const list = toolbarFonts({ disableBuiltin: ["Calibri"], fonts: [inter] });
    const values = list.map((f) => f.value);
    expect(values).not.toContain("Calibri, sans-serif"); // hidden
    expect(values).toContain("Arial, sans-serif"); // other built-ins stay
    expect(list.find((f) => f.value === "Inter")?.label).toBe("Inter"); // custom appended
    // disableBuiltin does NOT affect resolution — Calibri still maps to its clone.
    expect(cloneFamilyFor("Calibri")).toEqual({ clone: "Carlito", substituted: false });
  });

  it("toolbar hides custom defs that would fail registration (so they always resolve)", () => {
    vi.spyOn(console, "warn").mockImplementation(() => {});
    const woff = { ...inter, family: "BadWoff", faces: { regular: "https://x/B.woff2" } };
    const values = toolbarFonts({ fonts: [inter, woff] }).map((f) => f.value);
    expect(values).toContain("Inter");
    expect(values).not.toContain("BadWoff"); // cloneFamilyFor could never resolve it
  });

  it("registry instances are isolated; resolvers read the ACTIVE one", () => {
    const a = createFontRegistry();
    const b = createFontRegistry();
    a.register({ fonts: [inter] });
    b.register({ fonts: [{ ...inter, family: "Roboto", sizing: { ascent: 0.7, descent: 0.2 } }] });
    // Distinct synthetic face-key namespaces so two concurrent export jobs that both
    // define "Inter" with different bytes can't collide in the loaded map.
    expect(a.faceKey("Inter", "Regular")).not.toBe(b.faceKey("Inter", "Regular"));

    setActiveFontRegistry(a);
    expect(cloneFamilyFor("Inter")).toEqual({ clone: "Inter", substituted: false });
    expect(cloneFamilyFor("Roboto")).toEqual({ clone: "Arimo", substituted: true }); // not in A

    setActiveFontRegistry(b);
    expect(cloneFamilyFor("Roboto")).toEqual({ clone: "Roboto", substituted: false });
    expect(cloneFamilyFor("Inter")).toEqual({ clone: "Arimo", substituted: true }); // not in B
  });
});
