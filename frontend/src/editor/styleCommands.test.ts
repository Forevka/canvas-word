import { describe, expect, it } from "vitest";
import type { CharStyle, Document, NamedStyle, Paragraph, ParaStyle, SectionProps, Stylesheet } from "@cw/shared";
import { applyOp, descendantsOf, locateShape, mergeDuplicateNamedStyles, styleById, styleType, wouldCycle } from "@cw/shared";
import { bulletListDefinition, numberListDefinition } from "@cw/shared";
import { applyCharStyle, deleteListDefinition, deleteNamedStyle, mergeDuplicateStyles, renameNamedStyle, setDefaultStyle, upsertListDefinition, upsertNamedStyle, type NamedStyleSpec } from "./commands";
import type { Command, EditorState } from "./state";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 14, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA: ParaStyle = { align: "left", lineHeight: 1.35, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

let n = 0;
const p = (text: string, namedStyle?: string): Paragraph => ({
  kind: "paragraph", id: `p${n++}`, revision: 0,
  runs: [{ text, style: { ...CHAR } }],
  style: namedStyle ? { ...PARA, namedStyle } : { ...PARA },
});

const SHEET: Stylesheet = {
  defaultStyleId: "Normal",
  styles: [
    { id: "Normal", name: "Normal", type: "paragraph", char: { color: "#202124", fontSizePx: 16 }, para: { align: "left" } },
    { id: "Heading1", name: "Heading 1", type: "paragraph", basedOn: "Normal", char: { bold: true, fontSizePx: 24 }, para: { spaceBeforePx: 12 } },
  ],
};

const docOf = (sheet: Stylesheet, ...ps: Paragraph[]): Document => ({ section: SECTION, blocks: ps, stylesheet: sheet });

interface RunResult { doc: Document; inverses: ReturnType<typeof applyOp>["inverse"][] }
const run = (state: EditorState, cmd: Command): RunResult => {
  const trn = cmd(state);
  if (!trn) throw new Error("command produced no transaction");
  let doc = state.doc;
  const inverses: RunResult["inverses"] = [];
  for (const op of trn.ops) { const r = applyOp(doc, op); doc = r.doc; inverses.push(r.inverse); }
  return { doc, inverses };
};
const undo = ({ doc, inverses }: RunResult): Document => {
  let d = doc;
  for (let i = inverses.length - 1; i >= 0; i--) d = applyOp(d, inverses[i]!).doc;
  return d;
};
const sel = (b: Paragraph) => ({ anchor: { blockId: b.id, offset: 0 }, focus: { blockId: b.id, offset: 1 } });

describe("stylesheet pure helpers", () => {
  it("styleType defaults untyped styles to paragraph", () => {
    expect(styleType({ id: "X", name: "X", char: {}, para: {} } as NamedStyle)).toBe("paragraph");
    expect(styleType({ id: "X", name: "X", type: "character", char: {}, para: {} })).toBe("character");
  });
  it("descendantsOf collects the transitive children", () => {
    const sheet: Stylesheet = { defaultStyleId: "A", styles: [
      { id: "A", name: "A", char: {}, para: {} },
      { id: "B", name: "B", basedOn: "A", char: {}, para: {} },
      { id: "C", name: "C", basedOn: "B", char: {}, para: {} },
    ] };
    expect(descendantsOf(sheet, "A")).toEqual(new Set(["B", "C"]));
    expect(descendantsOf(sheet, "C")).toEqual(new Set());
  });
  it("wouldCycle rejects self and descendants", () => {
    expect(wouldCycle(SHEET, "Normal", "Heading1")).toBe(true); // Heading1 is Normal's child
    expect(wouldCycle(SHEET, "Heading1", "Heading1")).toBe(true);
    expect(wouldCycle(SHEET, "Heading1", "Normal")).toBe(false);
    expect(wouldCycle(SHEET, "Heading1", undefined)).toBe(false);
  });
});

describe("upsertNamedStyle", () => {
  it("creates a new paragraph style and applies it to the selection", () => {
    const a = p("hello");
    const state: EditorState = { doc: docOf(SHEET, a), selection: sel(a) };
    const spec: NamedStyleSpec = { name: "My Style", type: "paragraph", char: { color: "#ff0000" }, para: { align: "center" } };
    const { doc } = run(state, upsertNamedStyle(spec));
    const created = styleById(doc.stylesheet!, "MyStyle");
    expect(created).toBeTruthy();
    expect(created!.char.color).toBe("#ff0000");
    const para = doc.blocks[0] as Paragraph;
    expect(para.style.namedStyle).toBe("MyStyle");
    expect(para.style.align).toBe("center");
    expect(para.runs[0]!.style.color).toBe("#ff0000"); // baked onto the run
  });

  it("updating a style re-bakes every referencing paragraph", () => {
    const a = p("title", "Heading1");
    const b = p("body", "Normal");
    const state: EditorState = { doc: docOf(SHEET, a, b), selection: sel(a) };
    const spec: NamedStyleSpec = { id: "Heading1", name: "Heading 1", type: "paragraph", basedOn: "Normal", char: { bold: true, fontSizePx: 24, color: "#0000ff" }, para: { spaceBeforePx: 12 } };
    const { doc } = run(state, upsertNamedStyle(spec));
    const title = doc.blocks[0] as Paragraph;
    expect(title.runs[0]!.style.color).toBe("#0000ff"); // Heading1 paragraph re-baked
    const body = doc.blocks[1] as Paragraph;
    expect(body.runs[0]!.style.color).not.toBe("#0000ff"); // Normal paragraph untouched
  });

  it("rejects a based-on cycle (no-op)", () => {
    const a = p("x", "Normal");
    const state: EditorState = { doc: docOf(SHEET, a), selection: sel(a) };
    // Normal based on Heading1 (its own descendant) → cycle.
    const spec: NamedStyleSpec = { id: "Normal", name: "Normal", type: "paragraph", basedOn: "Heading1", char: {}, para: {} };
    expect(upsertNamedStyle(spec)({ ...state })).toBeNull();
  });

  it("create + undo restores the original stylesheet and paragraph", () => {
    const a = p("hello");
    const state: EditorState = { doc: docOf(SHEET, a), selection: sel(a) };
    const res = run(state, upsertNamedStyle({ name: "Temp", type: "paragraph", char: { color: "#abcdef" }, para: {} }));
    const back = undo(res);
    expect(styleById(back.stylesheet!, "Temp")).toBeUndefined();
    expect((back.blocks[0] as Paragraph).style.namedStyle).toBeUndefined();
  });
});

describe("applyCharStyle", () => {
  const charSheet: Stylesheet = { ...SHEET, styles: [...SHEET.styles, { id: "Emph", name: "Emphasis", type: "character", char: { italic: true, color: "#ff0000" }, para: {} }] };

  it("bakes the resolved char over the selected runs and tags charStyleId, leaving the paragraph alone", () => {
    const a = p("hello world", "Normal");
    const state: EditorState = { doc: docOf(charSheet, a), selection: { anchor: { blockId: a.id, offset: 0 }, focus: { blockId: a.id, offset: 5 } } };
    const { doc } = run(state, applyCharStyle("Emph"));
    const para = doc.blocks[0] as Paragraph;
    const first = para.runs[0]!;
    expect(first.style.charStyleId).toBe("Emph");
    expect(first.style.italic).toBe(true);
    expect(first.style.color).toBe("#ff0000");
    expect(para.style.namedStyle).toBe("Normal"); // paragraph style untouched
    // The unselected tail keeps no character style.
    expect(para.runs.some((r) => r.text.includes("world") && r.style.charStyleId)).toBe(false);
  });

  it("is a no-op on a collapsed caret or a paragraph style", () => {
    const a = p("x", "Normal");
    const collapsed: EditorState = { doc: docOf(charSheet, a), selection: { anchor: { blockId: a.id, offset: 0 }, focus: { blockId: a.id, offset: 0 } } };
    expect(applyCharStyle("Emph")(collapsed)).toBeNull();
    const ranged: EditorState = { doc: docOf(charSheet, a), selection: { anchor: { blockId: a.id, offset: 0 }, focus: { blockId: a.id, offset: 1 } } };
    expect(applyCharStyle("Heading1")(ranged)).toBeNull(); // not a character style
  });

  it("updating a character style re-bakes only its tagged runs", () => {
    const a = p("tagged", "Normal");
    // Tag the run first.
    const tagged = run({ doc: docOf(charSheet, a), selection: { anchor: { blockId: a.id, offset: 0 }, focus: { blockId: a.id, offset: 6 } } }, applyCharStyle("Emph")).doc;
    const state: EditorState = { doc: tagged, selection: null };
    const spec: NamedStyleSpec = { id: "Emph", name: "Emphasis", type: "character", char: { italic: true, color: "#00ff00" }, para: {} };
    const { doc } = run(state, upsertNamedStyle(spec));
    expect((doc.blocks[0] as Paragraph).runs[0]!.style.color).toBe("#00ff00");
  });
});

describe("mergeDuplicateNamedStyles (pure)", () => {
  it("collapses styles identical except name; the default survives", () => {
    const sheet: Stylesheet = { defaultStyleId: "Normal", styles: [
      { id: "Normal", name: "Normal", type: "paragraph", char: { color: "#111" }, para: { align: "left" } },
      { id: "BodyA", name: "Body A", type: "paragraph", char: { color: "#111" }, para: { align: "left" } },
      { id: "BodyB", name: "Body B", type: "paragraph", char: { color: "#111" }, para: { align: "left" } },
    ] };
    const { sheet: merged, remap } = mergeDuplicateNamedStyles(sheet);
    expect(merged.styles.map((s) => s.id)).toEqual(["Normal"]); // BodyA/BodyB folded into Normal
    expect(remap.get("BodyA")).toBe("Normal");
    expect(remap.get("BodyB")).toBe("Normal");
  });
  it("keeps genuinely different styles and rewrites basedOn through the remap", () => {
    const sheet: Stylesheet = { defaultStyleId: "Normal", styles: [
      { id: "Normal", name: "Normal", type: "paragraph", char: {}, para: {} },
      { id: "Blue", name: "Blue", type: "paragraph", char: { color: "#00f" }, para: {} },
      { id: "BlueDup", name: "Blue Dup", type: "paragraph", char: { color: "#00f" }, para: {} },
      { id: "Child", name: "Child", type: "paragraph", basedOn: "BlueDup", char: { bold: true }, para: {} },
    ] };
    const { sheet: merged, remap } = mergeDuplicateNamedStyles(sheet);
    expect(remap.get("BlueDup")).toBe("Blue");
    expect(styleById(merged, "Child")!.basedOn).toBe("Blue"); // rebased off the merged style
    expect(styleById(merged, "Blue")).toBeTruthy();
  });
  it("never merges across kinds (paragraph vs character)", () => {
    const sheet: Stylesheet = { defaultStyleId: "Normal", styles: [
      { id: "Normal", name: "Normal", type: "paragraph", char: { bold: true }, para: {} },
      { id: "Strong", name: "Strong", type: "character", char: { bold: true }, para: {} },
    ] };
    expect(mergeDuplicateNamedStyles(sheet).remap.size).toBe(0);
  });
});

describe("mergeDuplicateStyles command", () => {
  const dupSheet: Stylesheet = { defaultStyleId: "Normal", styles: [
    { id: "Normal", name: "Normal", type: "paragraph", char: { color: "#202124", fontSizePx: 16 }, para: { align: "left" } },
    { id: "Heading1", name: "Heading 1", type: "paragraph", basedOn: "Normal", char: { bold: true, fontSizePx: 24 }, para: { spaceBeforePx: 12 } },
    { id: "Heading1b", name: "Heading 1 (copy)", type: "paragraph", basedOn: "Normal", char: { bold: true, fontSizePx: 24 }, para: { spaceBeforePx: 12 } },
    { id: "Emph", name: "Emphasis", type: "character", char: { italic: true }, para: {} },
    { id: "Emph2", name: "Emphasis 2", type: "character", char: { italic: true }, para: {} },
  ] };

  it("remaps paragraph + run references to the survivor and drops the duplicates", () => {
    const a = p("title", "Heading1b");
    const b = p("body", "Normal");
    b.runs = [{ text: "body", style: { ...CHAR, italic: true, charStyleId: "Emph2" } }];
    const state: EditorState = { doc: docOf(dupSheet, a, b), selection: sel(a) };
    const { doc } = run(state, mergeDuplicateStyles());
    // Duplicates gone; survivors kept.
    expect(styleById(doc.stylesheet!, "Heading1b")).toBeUndefined();
    expect(styleById(doc.stylesheet!, "Emph2")).toBeUndefined();
    expect(styleById(doc.stylesheet!, "Heading1")).toBeTruthy();
    // References remapped.
    expect((doc.blocks[0] as Paragraph).style.namedStyle).toBe("Heading1");
    expect((doc.blocks[1] as Paragraph).runs[0]!.style.charStyleId).toBe("Emph");
  });

  it("is a no-op when there are no duplicates", () => {
    const a = p("x", "Normal");
    const sheet: Stylesheet = { ...SHEET };
    expect(mergeDuplicateStyles()({ doc: docOf(sheet, a), selection: sel(a) })).toBeNull();
  });

  it("undo restores the duplicate styles and references", () => {
    const a = p("title", "Heading1b");
    const res = run({ doc: docOf(dupSheet, a), selection: sel(a) }, mergeDuplicateStyles());
    const back = undo(res);
    expect(styleById(back.stylesheet!, "Heading1b")).toBeTruthy();
    expect((back.blocks[0] as Paragraph).style.namedStyle).toBe("Heading1b");
  });
});

describe("list definitions", () => {
  it("upserts a list definition (and undo restores)", () => {
    const a = p("x");
    const def = bulletListDefinition("custom-1", "◆");
    const res = run({ doc: docOf(SHEET, a), selection: sel(a) }, upsertListDefinition(def));
    expect(res.doc.lists!["custom-1"]!.levels[0]!.bulletChar).toBe("◆");
    const back = undo(res);
    expect(back.lists?.["custom-1"]).toBeUndefined();
  });

  it("replaces an existing definition and deletes one", () => {
    const a = p("x");
    let doc: Document = { ...docOf(SHEET, a), lists: { "custom-1": bulletListDefinition("custom-1", "•") } };
    doc = run({ doc, selection: null }, upsertListDefinition(numberListDefinition("custom-1", "decimal", "."))).doc;
    expect(doc.lists!["custom-1"]!.levels[0]!.format).toBe("decimal");
    doc = run({ doc, selection: null }, deleteListDefinition("custom-1")).doc;
    expect(doc.lists?.["custom-1"]).toBeUndefined();
  });
});

describe("deleteNamedStyle / rename / setDefault", () => {
  it("deletes a style, re-pointing referencing paragraphs to its base and re-baking", () => {
    const a = p("title", "Heading1");
    const state: EditorState = { doc: docOf(SHEET, a), selection: sel(a) };
    const { doc } = run(state, deleteNamedStyle("Heading1"));
    expect(styleById(doc.stylesheet!, "Heading1")).toBeUndefined();
    const para = doc.blocks[0] as Paragraph;
    expect(para.style.namedStyle).toBe("Normal"); // fell back to basedOn (same type)
  });

  it("cannot delete the default style", () => {
    const a = p("x", "Normal");
    const state: EditorState = { doc: docOf(SHEET, a), selection: sel(a) };
    expect(deleteNamedStyle("Normal")(state)).toBeNull();
  });

  it("reparents a child whose basedOn is deleted", () => {
    // Add a grandchild Heading2 → Heading1 → Normal; delete Heading1.
    const sheet: Stylesheet = { ...SHEET, styles: [...SHEET.styles, { id: "Heading2", name: "Heading 2", type: "paragraph", basedOn: "Heading1", char: {}, para: {} }] };
    const a = p("x", "Normal");
    const { doc } = run({ doc: docOf(sheet, a), selection: sel(a) }, deleteNamedStyle("Heading1"));
    expect(styleById(doc.stylesheet!, "Heading2")!.basedOn).toBe("Normal");
  });

  it("renames (name only, id preserved) and sets default", () => {
    const a = p("x", "Heading1");
    const renamed = run({ doc: docOf(SHEET, a), selection: null }, renameNamedStyle("Heading1", "Big Heading")).doc;
    expect(styleById(renamed.stylesheet!, "Heading1")!.name).toBe("Big Heading");
    const def = run({ doc: renamed, selection: null }, setDefaultStyle("Heading1")).doc;
    expect(def.stylesheet!.defaultStyleId).toBe("Heading1");
  });

  // Regression (#219): a text box nested INSIDE a table cell is read-only — its
  // paragraphs must NOT enter the editable paragraph space, or a paragraphsOf-
  // iterating editor (deleting a character style here, like find/replace) would
  // emit an op for a paragraph no content op can locate and throw.
  it("deleting a character style over a cell-nested text box does not throw (and leaves it untouched)", () => {
    const charSheet: Stylesheet = {
      defaultStyleId: "Normal",
      styles: [
        { id: "Normal", name: "Normal", type: "paragraph", char: {}, para: {} },
        { id: "Emph", name: "Emphasis", type: "character", char: { italic: true }, para: {} },
      ],
    };
    const withCs = (text: string, id: string): Paragraph => ({
      kind: "paragraph", id, revision: 0,
      runs: [{ text, style: { ...CHAR, charStyleId: "Emph" } }], style: { ...PARA },
    });
    // A text box nested in a table cell, whose run references the char style.
    const cellShape = {
      kind: "shape" as const, id: "cellbox", revision: 0, geometry: { preset: "rect" as const },
      widthPx: 200, heightPx: 90, align: "left" as const, text: { blocks: [withCs("boxed", "boxpar")] },
    };
    const table = {
      kind: "table" as const, id: "tbl", revision: 0,
      rows: [{ cells: [{ id: "c0", blocks: [p("cell"), cellShape] }] }],
    };
    const topPar = withCs("top-level", "toppar");
    const doc: Document = { section: SECTION, blocks: [topPar, table], stylesheet: charSheet };

    let result!: RunResult;
    expect(() => {
      result = run({ doc, selection: sel(topPar) }, deleteNamedStyle("Emph"));
    }).not.toThrow();

    // The TOP-LEVEL paragraph had its now-dangling char-style ref stripped…
    expect((result.doc.blocks[0] as Paragraph).runs[0]!.style.charStyleId).toBeUndefined();
    // …while the cell-nested text box run is left untouched (read-only sub-flow).
    const boxRun = ((result.doc.blocks[1] as { rows: { cells: { blocks: { text?: { blocks: Paragraph[] } }[] }[] }[] })
      .rows[0]!.cells[0]!.blocks[1] as { text: { blocks: Paragraph[] } }).text.blocks[0]!.runs[0]!;
    expect(boxRun.style.charStyleId).toBe("Emph");
  });

  // Regression (#219): Enter on a cell-nested text box is a no-op — the interactive
  // enter path gates on locateShape(...).kind === "top", so a cell-nested shape is
  // refused. This pins the location kind the gate reads.
  it("a text box inside a table cell is a cell-nested shape (Enter refuses it)", () => {
    const cellShape = {
      kind: "shape" as const, id: "cellbox", revision: 0, geometry: { preset: "rect" as const },
      widthPx: 200, heightPx: 90, align: "left" as const, text: { blocks: [p("boxed")] },
    };
    const topShape = {
      kind: "shape" as const, id: "topbox", revision: 0, geometry: { preset: "rect" as const },
      widthPx: 200, heightPx: 90, align: "left" as const, text: { blocks: [p("boxed")] },
    };
    const table = { kind: "table" as const, id: "tbl", revision: 0, rows: [{ cells: [{ id: "c0", blocks: [cellShape] }] }] };
    const doc: Document = { section: SECTION, blocks: [topShape, table], stylesheet: SHEET };
    expect(locateShape(doc, "cellbox")?.kind).toBe("cell"); // refused by the top-only gate
    expect(locateShape(doc, "topbox")?.kind).toBe("top"); // accepted
  });
});
