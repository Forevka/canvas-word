// Unit coverage for the custom-font overlay: resolution shadows built-ins, metrics
// come from the caller, the toolbar is filtered/extended per config, and invalid
// input (bad sizing, WOFF2) is rejected.

import { afterEach, describe, expect, it, vi } from "vitest";
import { cloneFamilyFor, metricsFor, toolbarFonts, CLONE_METRICS } from "./clones";
import {
  __resetCustomFonts,
  allCustomFonts,
  customFontFileName,
  customMetrics,
  registerCustomFonts,
  type CustomFontDef,
} from "./customRegistry";

const inter: CustomFontDef = {
  family: "Inter",
  faces: { regular: "https://x/Inter-Regular.ttf", bold: "https://x/Inter-Bold.ttf" },
  sizing: { ascent: 0.95, descent: 0.24 },
};

afterEach(() => {
  __resetCustomFonts();
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
    registerCustomFonts({ fonts: [{ ...inter, family: "Woff", faces: { regular: "https://x/F.woff2" } }] });
    expect(allCustomFonts().map((f) => f.family)).not.toContain("BadSize");
    expect(allCustomFonts().map((f) => f.family)).not.toContain("Woff");

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
});
