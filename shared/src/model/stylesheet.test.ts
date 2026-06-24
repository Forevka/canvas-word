import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, SectionProps } from "./document";
import {
  defaultStylesheet,
  descendantsOf,
  makeDefaultStylesheet,
  mergeDuplicateNamedStyles,
  type NamedStyle,
  paragraphsWithStyle,
  resolveCharStyle,
  resolveStyle,
  type Stylesheet,
  styleById,
  styleType,
  wouldCycle,
} from "./stylesheet";

const sheet = (styles: NamedStyle[], defaultStyleId = "Normal"): Stylesheet => ({ styles, defaultStyleId });
const style = (id: string, extra: Partial<NamedStyle> = {}): NamedStyle => ({ id, name: id, char: {}, para: {}, ...extra });

describe("styleType", () => {
  it("defaults untyped (legacy) styles to paragraph", () => {
    expect(styleType(style("Normal"))).toBe("paragraph");
  });

  it("returns the explicit type when present", () => {
    expect(styleType(style("Emph", { type: "character" }))).toBe("character");
  });
});

describe("styleById", () => {
  it("finds a style by id, undefined when absent", () => {
    const s = sheet([style("Normal"), style("Heading1")]);
    expect(styleById(s, "Heading1")?.id).toBe("Heading1");
    expect(styleById(s, "Nope")).toBeUndefined();
  });
});

describe("resolveStyle", () => {
  it("walks the basedOn chain root→leaf with leaf fields overriding", () => {
    const s = sheet([
      style("Normal", { char: { fontFamily: "Georgia", fontSizePx: 16, color: "#000" } }),
      style("Heading1", { basedOn: "Normal", char: { fontSizePx: 24, bold: true }, para: { spaceBeforePx: 18 } }),
    ]);
    const r = resolveStyle(s, "Heading1");
    expect(r.char).toEqual({ fontFamily: "Georgia", fontSizePx: 24, color: "#000", bold: true });
    expect(r.para).toEqual({ spaceBeforePx: 18 });
  });

  it("ignores undefined-valued fields (definedOnly) so they do not clobber the base", () => {
    // A child carrying an explicit `undefined` must NOT erase the inherited value.
    const childChar = { fontFamily: undefined, bold: true } as unknown as Partial<CharStyle>;
    const s = sheet([
      style("Normal", { char: { fontFamily: "Georgia" } }),
      style("Child", { basedOn: "Normal", char: childChar }),
    ]);
    expect(resolveStyle(s, "Child").char).toEqual({ fontFamily: "Georgia", bold: true });
  });

  it("returns empty deltas for an unknown id", () => {
    expect(resolveStyle(sheet([style("Normal")]), "Missing")).toEqual({ char: {}, para: {} });
  });

  it("terminates on a basedOn cycle instead of looping forever", () => {
    const s = sheet([
      style("A", { basedOn: "B", char: { bold: true } }),
      style("B", { basedOn: "A", char: { italic: true } }),
    ]);
    const r = resolveStyle(s, "A");
    expect(r.char).toEqual({ bold: true, italic: true });
  });
});

describe("resolveCharStyle", () => {
  it("resolves only the char chain for character styles", () => {
    const s = sheet([
      style("Normal", { char: { fontFamily: "Georgia", color: "#000" } }),
      style("Emph", { type: "character", basedOn: "Normal", char: { italic: true, color: "#555" } }),
    ]);
    expect(resolveCharStyle(s, "Emph")).toEqual({ fontFamily: "Georgia", italic: true, color: "#555" });
  });
});

describe("descendantsOf", () => {
  it("returns every transitive descendant of a style", () => {
    const s = sheet([
      style("Normal"),
      style("Heading1", { basedOn: "Normal" }),
      style("Heading2", { basedOn: "Heading1" }),
      style("Quote", { basedOn: "Normal" }),
    ]);
    expect(descendantsOf(s, "Normal")).toEqual(new Set(["Heading1", "Heading2", "Quote"]));
    expect(descendantsOf(s, "Heading1")).toEqual(new Set(["Heading2"]));
    expect(descendantsOf(s, "Quote")).toEqual(new Set());
  });
});

describe("wouldCycle", () => {
  const s = sheet([
    style("Normal"),
    style("Heading1", { basedOn: "Normal" }),
    style("Heading2", { basedOn: "Heading1" }),
  ]);

  it("is false for undefined newBasedOn", () => {
    expect(wouldCycle(s, "Heading1", undefined)).toBe(false);
  });

  it("is true for a self-reference", () => {
    expect(wouldCycle(s, "Heading1", "Heading1")).toBe(true);
  });

  it("is true when newBasedOn is a descendant of the style", () => {
    expect(wouldCycle(s, "Heading1", "Heading2")).toBe(true);
  });

  it("is false for a non-descendant target", () => {
    expect(wouldCycle(s, "Heading1", "Normal")).toBe(false);
  });
});

describe("mergeDuplicateNamedStyles", () => {
  it("returns the sheet unchanged with empty remap when there are no duplicates", () => {
    const s = sheet([style("Normal", { char: { bold: true } }), style("Other", { char: { italic: true } })]);
    const r = mergeDuplicateNamedStyles(s);
    expect(r.remap.size).toBe(0);
    expect(r.sheet).toBe(s);
  });

  it("collapses same-signature styles, keeping the first as survivor", () => {
    const s = sheet([
      style("Normal", { char: { fontSizePx: 16 } }),
      style("Dup1", { name: "Dup1", char: { italic: true } }),
      style("Dup2", { name: "Dup2-different-name", char: { italic: true } }),
    ]);
    const r = mergeDuplicateNamedStyles(s);
    expect(r.remap.get("Dup2")).toBe("Dup1");
    expect(r.sheet.styles.map((x) => x.id)).toEqual(["Normal", "Dup1"]);
  });

  it("promotes the default style as survivor even if it appears later", () => {
    const s = sheet(
      [
        style("First", { char: { italic: true } }),
        style("Normal", { char: { italic: true } }),
      ],
      "Normal",
    );
    const r = mergeDuplicateNamedStyles(s);
    expect(r.remap.get("First")).toBe("Normal");
    expect(r.sheet.defaultStyleId).toBe("Normal");
    expect(r.sheet.styles.map((x) => x.id)).toEqual(["Normal"]);
  });

  it("rewrites basedOn through the remap and drops self-references after merge", () => {
    const s = sheet([
      style("Normal", { char: { fontSizePx: 16 } }),
      style("A", { char: { bold: true } }),
      style("B", { char: { bold: true } }), // dup of A → remapped to A
      style("Child", { basedOn: "B", char: { italic: true } }),
    ]);
    const r = mergeDuplicateNamedStyles(s);
    const child = r.sheet.styles.find((x) => x.id === "Child")!;
    expect(child.basedOn).toBe("A");
  });
});

describe("paragraphsWithStyle", () => {
  const section = (): SectionProps => ({ pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } });
  const CHAR: CharStyle = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
  const para = (id: string, namedStyle?: string): Paragraph => ({
    kind: "paragraph",
    id,
    revision: 0,
    runs: [{ text: id, style: CHAR }],
    style: { align: "left", lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0, ...(namedStyle ? { namedStyle } : {}) },
  });

  it("returns paragraphs referencing a given style id, in document order", () => {
    const doc: Document = { section: section(), blocks: [para("a", "Heading1"), para("b", "Normal"), para("c", "Heading1")] };
    expect(paragraphsWithStyle(doc, "Heading1").map((p) => p.id)).toEqual(["a", "c"]);
  });

  it("returns empty array when no paragraph references the style", () => {
    const doc: Document = { section: section(), blocks: [para("a", "Normal")] };
    expect(paragraphsWithStyle(doc, "Quote")).toEqual([]);
  });
});

describe("makeDefaultStylesheet / defaultStylesheet", () => {
  it("produces the Normal-defaulted built-in gallery with no overrides", () => {
    const s = defaultStylesheet();
    expect(s.defaultStyleId).toBe("Normal");
    expect(s.styles.map((x) => x.id)).toEqual(["Normal", "Title", "Subtitle", "Heading1", "Heading2", "Quote", "Code"]);
    expect(styleById(s, "Normal")!.char).toMatchObject({ fontFamily: "Georgia, serif", fontSizePx: 16, color: "#202124" });
  });

  it("applies typography overrides to Normal + heading font", () => {
    const s = makeDefaultStylesheet({ fontFamily: "Times, serif", fontSizePx: 18, color: "#111", headingFontFamily: "Helvetica", lineHeight: 2 });
    expect(styleById(s, "Normal")!.char).toMatchObject({ fontFamily: "Times, serif", fontSizePx: 18, color: "#111" });
    expect(styleById(s, "Normal")!.para.lineHeight).toBe(2);
    expect(styleById(s, "Heading1")!.char.fontFamily).toBe("Helvetica");
  });

  it("zero-arg matches the explicit no-override result", () => {
    expect(defaultStylesheet()).toEqual(makeDefaultStylesheet());
  });
});
