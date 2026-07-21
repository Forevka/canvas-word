// Extended builder surface: equations (block + inline, LaTeX + MathML), RTL /
// writing direction, outline level, keep-together, tab stops, character-style
// application, cell preferred width, table autofit, and REAL table styles
// (registered entity + baked cells + round-trip). Model-level assertions plus one
// end-to-end docx round-trip. Pure Node except the round-trip (needs measure host).

import { beforeAll, describe, expect, it } from "vitest";
import type { EquationBlock, Paragraph, TableBlock } from "@cw/shared";
import { DocumentBuilder } from "./index";

const para = (b: unknown): Paragraph => b as Paragraph;
const tableOf = (b: unknown): TableBlock => b as TableBlock;
const FRAC_ML = "<math><mfrac><mi>a</mi><mn>2</mn></mfrac></math>";

describe("equations", () => {
  it("equation(latex) appends a centered display EquationBlock", () => {
    const doc = DocumentBuilder.create({ idSeed: "eq" }).equation("\\frac{a}{2}", { align: "right" }).build();
    const eq = doc.blocks.find((b): b is EquationBlock => b.kind === "equation")!;
    expect(eq.equation.display).toBe(true);
    expect(eq.align).toBe("right");
    expect(eq.equation.root.children.some((n) => n.type === "frac")).toBe(true);
  });

  it("equation(latex, {scale}) sets a clamped uniform scale; scale 1 is omitted", () => {
    const doc = DocumentBuilder.create().equation("x", { scale: 1.5 }).equation("y", { scale: 99 }).equation("z", { scale: 1 }).build();
    const eqs = doc.blocks.filter((b): b is EquationBlock => b.kind === "equation");
    expect(eqs[0]!.scale).toBe(1.5);
    expect(eqs[1]!.scale).toBe(4); // clamped to the max
    expect(eqs[2]!.scale).toBeUndefined(); // default size carries no field
  });

  it("equationMathml builds the same AST from a MathML string", () => {
    const doc = DocumentBuilder.create().equationMathml(FRAC_ML).build();
    const eq = doc.blocks.find((b): b is EquationBlock => b.kind === "equation")!;
    expect(eq.equation.display).toBe(true);
    expect(eq.equation.root.children.some((n) => n.type === "frac")).toBe(true);
  });

  it("inlineEquation appends a U+FFFC run carrying an inline (display:false) equation", () => {
    const doc = DocumentBuilder.create()
      .paragraph("the identity ")
      .inlineEquation("e^{i\\pi}+1=0")
      .text(".")
      .build();
    const runs = para(doc.blocks[0]).runs;
    const eqRun = runs.find((r) => r.style.equation)!;
    expect(eqRun.text).toBe("￼");
    expect(eqRun.style.equation!.display).toBe(false);
    expect(runs[runs.length - 1]!.text).toBe(".");
  });
});

describe("writing direction + paragraph flags", () => {
  it("direction('rtl') sets the paragraph base direction", () => {
    const doc = DocumentBuilder.create().paragraph("شلام").direction("rtl").build();
    expect(para(doc.blocks[0]).style.direction).toBe("rtl");
  });

  it("rtl() sets the per-run w:rtl flag on existing + later runs", () => {
    const doc = DocumentBuilder.create().paragraph("a").rtl().text("b").build();
    const runs = para(doc.blocks[0]).runs;
    expect(runs.every((r) => r.style.rtl === true)).toBe(true);
  });

  it("keepTogether + outlineLevel (clamped) set their ParaStyle fields", () => {
    const doc = DocumentBuilder.create().paragraph("h").keepTogether().outlineLevel(99).build();
    expect(para(doc.blocks[0]).style.keepLinesTogether).toBe(true);
    expect(para(doc.blocks[0]).style.outlineLevel).toBe(8);
  });

  it("contextualSpacing() sets the ParaStyle flag (and clears with false)", () => {
    const on = DocumentBuilder.create().paragraph("v").contextualSpacing().build();
    expect(para(on.blocks[0]).style.contextualSpacing).toBe(true);
    const off = DocumentBuilder.create().paragraph("v").contextualSpacing(false).build();
    expect(para(off.blocks[0]).style.contextualSpacing).toBe(false);
  });

  it("tabStops are stored sorted by position", () => {
    const doc = DocumentBuilder.create()
      .paragraph("x")
      .tabStops([{ posPx: 300, align: "right" }, { posPx: 100, leader: "dot" }])
      .build();
    const stops = para(doc.blocks[0]).style.tabStops!;
    expect(stops.map((s) => s.posPx)).toEqual([100, 300]);
  });
});

describe("character styles", () => {
  it("charStyle bakes the style's char AND sets the w:rStyle reference", () => {
    const doc = DocumentBuilder.create()
      .style({ id: "Code", name: "Code", type: "character", char: { fontFamily: "Consolas, monospace", color: "#a31515" }, para: {} })
      .paragraph("x = 1")
      .charStyle("Code")
      .build();
    const run = para(doc.blocks[0]).runs[0]!;
    expect(run.style.charStyleId).toBe("Code");
    expect(run.style.fontFamily).toBe("Consolas, monospace");
    expect(run.style.color).toBe("#a31515");
  });

  it("charStyle rejects a PARAGRAPH style (w:rStyle must be a character style)", () => {
    const builder = DocumentBuilder.create()
      .style({ id: "Body", name: "Body", type: "paragraph", char: { color: "#111" }, para: {} });
    builder.paragraph("x").charStyle("Body");
    const doc = builder.build();
    expect(para(doc.blocks[0]).runs[0]!.style.charStyleId).toBeUndefined();
    expect(builder.warnings.some((w) => w.code.startsWith("char-style-invalid"))).toBe(true);
  });
});

describe("table cell width + autofit", () => {
  it("preferredWidth + widthMode reach the model", () => {
    const doc = DocumentBuilder.create()
      .table([[{ text: "A", preferredWidth: { px: 120, type: "abs" } }, "B"]], { widthMode: "autofitContents" })
      .build();
    const t = tableOf(doc.blocks[0]);
    expect(t.widthMode).toBe("autofitContents");
    expect(t.rows[0]!.cells[0]!.preferredWidth).toEqual({ px: 120, type: "abs" });
  });
});

describe("real table styles", () => {
  const gridStyle = {
    id: "GridAccent",
    name: "Grid Accent",
    conds: {
      wholeTable: { borders: { top: { color: "#888888", widthPx: 1 }, bottom: { color: "#888888", widthPx: 1 }, left: { color: "#888888", widthPx: 1 }, right: { color: "#888888", widthPx: 1 } } },
      firstRow: { shading: "#1a1a2e", char: { bold: true, color: "#ffffff" } },
      band2Horz: { shading: "#f1f3f4" },
    },
  } as const;

  it("registers the style and bakes shading onto the header + sets the reference", () => {
    const doc = DocumentBuilder.create({ idSeed: "ts" })
      .tableStyle(gridStyle)
      .table([["H1", "H2"], ["a", "b"], ["c", "d"]], { styleId: "GridAccent" })
      .build();
    expect(doc.tableStyles!.GridAccent).toBeTruthy();
    const t = tableOf(doc.blocks[0]);
    expect(t.styleId).toBe("GridAccent");
    expect(t.condOverrides).toEqual({ firstRow: true, bandRows: true });
    // Header row shading is baked (firstRow band active by default tblLook).
    expect(t.rows[0]!.cells[0]!.shading).toBe("#1a1a2e");
    // Banded body row (row index 2, second body row) gets band2Horz shading.
    expect(t.rows[2]!.cells[0]!.shading).toBe("#f1f3f4");
    // The header band's char (bold + white) is baked onto the cell content so it
    // renders headlessly, not just shipped in w:tblStyle.
    const headerRun = para(t.rows[0]!.cells[0]!.blocks[0]).runs[0]!;
    expect(headerRun.style.bold).toBe(true);
    expect(headerRun.style.color).toBe("#ffffff");
  });

  it("preserves explicit cell content formatting over the band, but defaults pick it up", () => {
    const doc = DocumentBuilder.create({ idSeed: "tp" })
      .tableStyle(gridStyle)
      .table([[{ text: "Red", style: { color: "#ff0000" } }, "Plain"]], { styleId: "GridAccent" })
      .build();
    const t = tableOf(doc.blocks[0]);
    const explicit = para(t.rows[0]!.cells[0]!.blocks[0]).runs[0]!;
    const plain = para(t.rows[0]!.cells[1]!.blocks[0]).runs[0]!;
    expect(explicit.style.color).toBe("#ff0000"); // explicit color kept over the band's white
    expect(explicit.style.bold).toBe(true); // bold was default → band's bold applies
    expect(plain.style.color).toBe("#ffffff"); // default color → band's white applies
  });

  it("style (preset) and styleId (real) are mutually exclusive — styleId wins with a warning", () => {
    const builder = DocumentBuilder.create({ idSeed: "tx" })
      .tableStyle(gridStyle)
      .table([["H"], ["x"]], { style: "headerBand", styleId: "GridAccent" });
    const doc = builder.build();
    expect(tableOf(doc.blocks[0]).styleId).toBe("GridAccent");
    expect(builder.warnings.some((w) => w.code === "table-style-conflict")).toBe(true);
  });

  it("an unregistered styleId sets the reference but warns and bakes nothing", () => {
    const builder = DocumentBuilder.create().table([["x"]], { styleId: "Missing" });
    const doc = builder.build();
    expect(tableOf(doc.blocks[0]).styleId).toBe("Missing");
    expect(builder.warnings.some((w) => w.code.startsWith("table-style-ref-missing"))).toBe(true);
  });
});

describe("block-level content controls", () => {
  it("contentControlRange wraps every appended block in a block-level SDT", () => {
    const builder = DocumentBuilder.create({ idSeed: "cc" })
      .contentControlRange({ type: "richText", alias: "Editable", tag: "region1" }, (s) => {
        s.paragraph("inside one");
        s.paragraph("inside two");
      })
      .paragraph("outside");
    const doc = builder.build();
    const inside = [para(doc.blocks[0]), para(doc.blocks[1])];
    const sid = inside[0]!.sdtPath![0]!;
    expect(inside.every((p) => p.sdtPath?.[0] === sid)).toBe(true);
    expect(doc.sdts![sid]).toEqual({ type: "richText", alias: "Editable", tag: "region1" });
    // The block after the range is not wrapped.
    expect(para(doc.blocks[2]).sdtPath).toBeUndefined();
  });

  it("nested ranges compose outer→inner on the shared blocks", () => {
    const doc = DocumentBuilder.create({ idSeed: "cc2" })
      .contentControlRange({ type: "richText", tag: "outer" }, (s) => {
        s.contentControlRange({ type: "richText", tag: "inner" }, (s2) => s2.paragraph("deep"));
      })
      .build();
    const path = para(doc.blocks[0]).sdtPath!;
    expect(path.length).toBe(2);
    expect(doc.sdts![path[0]!]!.tag).toBe("outer");
    expect(doc.sdts![path[1]!]!.tag).toBe("inner");
  });
});

describe("docx round-trip (built doc)", () => {
  beforeAll(async () => {
    const { installMeasureHost } = await import("../export/shared/measureHost");
    await installMeasureHost();
  });

  it("a built block equation + real table style survive export → import", async () => {
    const { runExport } = await import("../export/pipeline");
    const { runImport } = await import("../import/docx/pipeline");
    const doc = DocumentBuilder.create({ idSeed: "rt" })
      .tableStyle({ id: "RT", name: "RT", conds: { firstRow: { shading: "#1a1a2e" } } })
      .equation("\\frac{a}{2}")
      .table([["H"], ["x"]], { styleId: "RT" })
      .build();

    const { bytes } = await runExport(doc, "docx");
    const { doc: back } = runImport(bytes);

    const eq = back.blocks.find((b): b is EquationBlock => b.kind === "equation");
    expect(eq?.equation.root.children.some((n) => n.type === "frac")).toBe(true);
    const t = back.blocks.find((b): b is TableBlock => b.kind === "table")!;
    expect(t.styleId).toBe("RT");
    expect(back.tableStyles!.RT).toBeTruthy();
  });
});
