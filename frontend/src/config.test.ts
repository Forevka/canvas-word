import { describe, it, expect } from "vitest";
import { DEFAULT_BEHAVIOR, DEFAULT_THEME, darkCanvasTheme, resolveBehavior, resolveConfig, resolveTheme } from "./config";

describe("resolveTheme", () => {
  it("returns the built-in default when no override is given", () => {
    expect(resolveTheme()).toEqual(DEFAULT_THEME);
  });

  it("merges a partial over the default, keeping unspecified fields", () => {
    const t = resolveTheme({ grid: "#ff0000", canvasBackground: "#202124" });
    expect(t.grid).toBe("#ff0000");
    expect(t.canvasBackground).toBe("#202124");
    // Untouched fields keep the built-in value.
    expect(t.caret).toBe(DEFAULT_THEME.caret);
    expect(t.tocLeader).toBe(DEFAULT_THEME.tocLeader);
  });

  it("deep-merges the nested ruler object", () => {
    const t = resolveTheme({ ruler: { bg: "#000000" } });
    expect(t.ruler.bg).toBe("#000000");
    // Sibling ruler fields stay at the default.
    expect(t.ruler.content).toBe(DEFAULT_THEME.ruler.content);
    expect(t.ruler.font).toBe(DEFAULT_THEME.ruler.font);
  });

  it("resolves the exported darkCanvasTheme preset", () => {
    const t = resolveTheme(darkCanvasTheme);
    expect(t.canvasBackground).toBe("#202124");
    expect(t.ruler.bg).toBe("#2a2d31");
  });
});

describe("resolveBehavior", () => {
  it("defaults then overrides", () => {
    expect(resolveBehavior()).toEqual(DEFAULT_BEHAVIOR);
    const b = resolveBehavior({ zoomStep: 1.25 });
    expect(b.zoomStep).toBe(1.25);
    expect(b.zoomMin).toBe(DEFAULT_BEHAVIOR.zoomMin);
    expect(b.indentStepPx).toBe(DEFAULT_BEHAVIOR.indentStepPx);
  });
});

describe("resolveConfig", () => {
  it("builds a default stylesheet from the typography override", () => {
    const cfg = resolveConfig({ overrideDefaultStyles: { fontFamily: "Times New Roman, serif", fontSizePx: 14 } });
    const normal = cfg.stylesheet.styles.find((s) => s.id === "Normal")!;
    expect(normal.char.fontFamily).toBe("Times New Roman, serif");
    expect(normal.char.fontSizePx).toBe(14);
    expect(cfg.typography.fontFamily).toBe("Times New Roman, serif");
  });

  it("with no input reproduces the built-in look", () => {
    const cfg = resolveConfig();
    expect(cfg.theme).toEqual(DEFAULT_THEME);
    expect(cfg.behavior).toEqual(DEFAULT_BEHAVIOR);
    expect(cfg.stylesheet.styles.find((s) => s.id === "Normal")!.char.fontFamily).toBe("Georgia, serif");
  });

  it("defaults develop mode off and honors the override", () => {
    expect(resolveConfig().develop).toBe(false);
    expect(resolveConfig({ develop: false }).develop).toBe(false);
    expect(resolveConfig({ develop: true }).develop).toBe(true);
  });
});
