// The setEquation / setEquationAlign ops: replace a display equation's MathML or
// alignment, exactly invertible (undo restores the prior value). Both are
// container-aware — the equation may live in the body, a header/footer band, or a
// table cell — and resolve through the generic locateBlock/locateEquation helper.

import { describe, expect, it } from "vitest";
import { applyOp, locateEquation, locateBlock } from "./ops";
import type { Block, Document, EquationBlock, Paragraph, TableBlock } from "./document";
import type { MathEquation } from "./math";

/** A minimal single-number display equation (just enough to assert against). */
const eq = (text: string): MathEquation => ({ root: { type: "row", children: [{ type: "number", text }] }, display: true });

/** A center-aligned display-equation block carrying the given equation. */
const equationBlock = (id: string, equation: MathEquation): EquationBlock => ({
  kind: "equation",
  id,
  revision: 0,
  equation,
  align: "center",
});

/** An empty paragraph block (used as a sibling/neighbour in fixtures). */
const para = (id: string): Paragraph => ({
  kind: "paragraph",
  id,
  revision: 0,
  runs: [{ text: "", style: { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" } }],
  style: { align: "left", lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 },
});

/** A 1x1 table whose single cell holds the given blocks. */
const tableWith = (id: string, blocks: Block[]): TableBlock => ({
  kind: "table",
  id,
  revision: 0,
  rows: [{ cells: [{ id: `${id}-c0`, blocks }] }],
});

/** A document with the given body blocks and optional section overrides (bands). */
const docOf = (blocks: Block[], section: Partial<Document["section"]> = {}): Document => ({
  section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 }, ...section },
  blocks,
});

/** A document whose sole body block is a center-aligned equation (id "e1"). */
const docWith = (equation: MathEquation): Document => docOf([equationBlock("e1", equation)]);

describe("setEquation op", () => {
  it("replaces the equation and inverts cleanly", () => {
    const doc = docWith(eq("1"));
    const r = applyOp(doc, { type: "setEquation", blockId: "e1", equation: eq("2") });
    const after = r.doc.blocks[0] as EquationBlock;
    expect(after.equation.root.children[0]).toMatchObject({ type: "number", text: "2" });
    expect(after.revision).toBe(1);

    // Inverse restores the original.
    const back = applyOp(r.doc, r.inverse).doc.blocks[0] as EquationBlock;
    expect(back.equation.root.children[0]).toMatchObject({ type: "number", text: "1" });
  });

  it("throws for a non-equation block id", () => {
    const doc = docWith(eq("1"));
    expect(() => applyOp(doc, { type: "setEquation", blockId: "missing", equation: eq("9") })).toThrow();
  });

  it("throws when the id is a top-level block of another kind", () => {
    const doc = docOf([para("p1")]);
    expect(() => applyOp(doc, { type: "setEquation", blockId: "p1", equation: eq("9") })).toThrow();
  });
});

describe("setEquationAlign op", () => {
  it("sets the alignment and inverts to the prior value", () => {
    const doc = docWith(eq("1")); // default align "center"
    const r = applyOp(doc, { type: "setEquationAlign", blockId: "e1", align: "right" });
    expect((r.doc.blocks[0] as EquationBlock).align).toBe("right");
    const back = applyOp(r.doc, r.inverse).doc.blocks[0] as EquationBlock;
    expect(back.align).toBe("center");
  });
});

describe("nested equations (table cells / header-footer bands)", () => {
  /** The first block of the first body table's single cell, as an equation. */
  const cellEquationOf = (doc: Document): EquationBlock =>
    (doc.blocks[0] as TableBlock).rows[0]!.cells[0]!.blocks[0] as EquationBlock;

  it("edits an equation inside a table cell and inverts", () => {
    const doc = docOf([tableWith("t1", [equationBlock("e1", eq("1"))])]);
    const r = applyOp(doc, { type: "setEquation", blockId: "e1", equation: eq("2") });
    expect(cellEquationOf(r.doc).equation.root.children[0]).toMatchObject({ type: "number", text: "2" });
    expect(cellEquationOf(r.doc).revision).toBe(1);
    // The table itself is path-cloned (revision bumped) so layout re-runs.
    expect((r.doc.blocks[0] as TableBlock).revision).toBe(1);

    const back = applyOp(r.doc, r.inverse).doc;
    expect(cellEquationOf(back).equation.root.children[0]).toMatchObject({ type: "number", text: "1" });
  });

  it("aligns an equation inside a table cell and inverts", () => {
    const doc = docOf([tableWith("t1", [para("p0"), equationBlock("e1", eq("1"))])]);
    const r = applyOp(doc, { type: "setEquationAlign", blockId: "e1", align: "left" });
    const cell = (r.doc.blocks[0] as TableBlock).rows[0]!.cells[0]!;
    expect((cell.blocks[1] as EquationBlock).align).toBe("left");
    // Sibling paragraph is untouched.
    expect(cell.blocks[0]!.id).toBe("p0");

    const back = applyOp(r.doc, r.inverse).doc;
    const backCell = (back.blocks[0] as TableBlock).rows[0]!.cells[0]!;
    expect((backCell.blocks[1] as EquationBlock).align).toBe("center");
  });

  it("edits an equation in a header band and inverts", () => {
    const doc = docOf([para("body0")], { header: [equationBlock("e1", eq("1"))] });
    const r = applyOp(doc, { type: "setEquation", blockId: "e1", equation: eq("2") });
    expect((r.doc.section.header![0] as EquationBlock).equation.root.children[0]).toMatchObject({ type: "number", text: "2" });
    // The body story is left alone.
    expect(r.doc.blocks[0]!.id).toBe("body0");

    const back = applyOp(r.doc, r.inverse).doc;
    expect((back.section.header![0] as EquationBlock).equation.root.children[0]).toMatchObject({ type: "number", text: "1" });
  });

  it("edits an equation inside a table cell within a footer band", () => {
    const doc = docOf([para("body0")], { footer: [tableWith("ft", [equationBlock("e1", eq("1"))])] });
    const r = applyOp(doc, { type: "setEquationAlign", blockId: "e1", align: "right" });
    const ftCell = (r.doc.section.footer![0] as TableBlock).rows[0]!.cells[0]!;
    expect((ftCell.blocks[0] as EquationBlock).align).toBe("right");
  });
});

describe("locateEquation / locateBlock", () => {
  it("locates a body top-level equation", () => {
    const doc = docWith(eq("1"));
    const loc = locateEquation(doc, "e1");
    expect(loc).toMatchObject({ kind: "top", where: "body", index: 0 });
  });

  it("locates an equation inside a table cell with its full path", () => {
    const doc = docOf([para("p0"), tableWith("t1", [para("c0"), equationBlock("e1", eq("1"))])]);
    const loc = locateEquation(doc, "e1");
    expect(loc).toMatchObject({ kind: "cell", where: "body", bi: 1, ri: 0, ci: 0, ii: 1 });
  });

  it("locates an equation in a band", () => {
    const doc = docOf([para("body0")], { header: [para("h0"), equationBlock("e1", eq("1"))] });
    const loc = locateEquation(doc, "e1");
    expect(loc).toMatchObject({ kind: "top", where: "header", index: 1 });
  });

  it("returns null when the id is not an equation, and respects the requested kind", () => {
    const doc = docOf([tableWith("t1", [equationBlock("e1", eq("1"))])]);
    expect(locateEquation(doc, "missing")).toBeNull();
    expect(locateEquation(doc, "t1")).toBeNull(); // a table, not an equation
    // The same locator finds the table when asked for that kind.
    expect(locateBlock(doc, "t1", "table")).toMatchObject({ kind: "top", where: "body", index: 0 });
  });
});
