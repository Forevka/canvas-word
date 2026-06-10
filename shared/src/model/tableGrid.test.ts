import { describe, expect, it } from "vitest";
import type { TableBlock, TableCell } from "./document";
import { buildTableGrid, cellsInRect, gridOriginOfCell, mergeRows, normalizeRect, unmergeRows } from "./tableGrid";

let n = 0;
const cell = (label: string, span?: { col?: number; row?: number }): TableCell => ({
  id: `c${n++}`,
  blocks: [{ kind: "paragraph", id: `p${n++}`, revision: 0, runs: [{ text: label, style: {} as never }], style: {} as never }],
  ...(span?.col ? { colSpan: span.col } : {}),
  ...(span?.row ? { rowSpan: span.row } : {}),
});
const table = (rows: TableCell[][]): TableBlock => ({
  kind: "table",
  id: "t",
  revision: 0,
  rows: rows.map((cells) => ({ cells })),
});

describe("buildTableGrid", () => {
  it("maps a plain grid 1:1", () => {
    const t = table([
      [cell("a"), cell("b")],
      [cell("c"), cell("d")],
    ]);
    const g = buildTableGrid(t);
    expect([g.rows, g.cols]).toEqual([2, 2]);
    expect(g.slots[0]![0]!.cell.blocks).toBeDefined();
    expect(g.slots[1]![1]!.ci).toBe(1);
  });

  it("fills a colSpan across its columns", () => {
    const t = table([[cell("wide", { col: 2 })], [cell("c"), cell("d")]]);
    const g = buildTableGrid(t);
    expect(g.cols).toBe(2);
    expect(g.slots[0]![0]).toBe(g.slots[0]![1]); // same slot object spans both cols
    expect(g.slots[0]![0]!.colSpan).toBe(2);
  });

  it("fills a rowSpan down its rows, shifting later cells past the hole", () => {
    // row0: [A(rowSpan2), B] ; row1: [C] — C must land in column 1, not 0.
    const t = table([[cell("A", { row: 2 }), cell("B")], [cell("C")]]);
    const g = buildTableGrid(t);
    expect(g.cols).toBe(2);
    expect(g.slots[1]![0]!.cell).toBe(g.slots[0]![0]!.cell); // A occupies (1,0)
    expect(g.slots[1]![1]!.cell.blocks[0]).toBeDefined(); // C at (1,1)
    expect(g.slots[1]![1]!.originRow).toBe(1);
  });
});

describe("normalizeRect", () => {
  it("expands to cover a straddling merged cell", () => {
    // A spans cols 0-1 on row 0. Selecting just (0,0) must pull in (0,1).
    const t = table([[cell("A", { col: 2 })], [cell("c"), cell("d")]]);
    const g = buildTableGrid(t);
    expect(normalizeRect(g, { r0: 0, c0: 0, r1: 0, c1: 0 })).toEqual({ r0: 0, c0: 0, r1: 0, c1: 1 });
  });

  it("orders inverted corners and clamps to bounds", () => {
    const t = table([[cell("a"), cell("b")], [cell("c"), cell("d")]]);
    const g = buildTableGrid(t);
    expect(normalizeRect(g, { r0: 1, c0: 1, r1: 0, c1: 0 })).toEqual({ r0: 0, c0: 0, r1: 1, c1: 1 });
  });
});

describe("cellsInRect", () => {
  it("returns distinct owners once, in reading order", () => {
    const t = table([[cell("A", { col: 2 })], [cell("c"), cell("d")]]);
    const g = buildTableGrid(t);
    const got = cellsInRect(g, { r0: 0, c0: 0, r1: 1, c1: 1 }).map((s) => s.cell.blocks[0]);
    expect(got.length).toBe(3); // A once, plus c and d
  });
});

const merged = (id: string, span: { col: number; row: number }): TableCell => ({
  id,
  blocks: [{ kind: "paragraph", id: "pm", revision: 0, runs: [{ text: "m", style: {} as never }], style: {} as never }],
  colSpan: span.col,
  rowSpan: span.row,
});

describe("mergeRows", () => {
  it("collapses a 2x2 block into one rowSpan+colSpan cell, emptying covered rows", () => {
    const t = table([
      [cell("a"), cell("b"), cell("x")],
      [cell("c"), cell("d"), cell("y")],
    ]);
    const g = buildTableGrid(t);
    const rect = normalizeRect(g, { r0: 0, c0: 0, r1: 1, c1: 1 });
    const rows = mergeRows(g, rect, merged("M", { col: 2, row: 2 }));
    // row0: [M(2x2), x] ; row1: [y] (cols 0-1 continue from M).
    expect(rows[0]!.cells.map((c) => c.id)).toEqual(["M", t.rows[0]!.cells[2]!.id]);
    expect(rows[1]!.cells.map((c) => c.id)).toEqual([t.rows[1]!.cells[2]!.id]);
    // Re-grid the merged result: M must own all four top-left slots.
    const g2 = buildTableGrid({ ...t, rows });
    expect(g2.slots[1]![1]!.cell.id).toBe("M");
    expect(g2.slots[0]![2]!.cell.id).toBe(t.rows[0]!.cells[2]!.id);
  });

  it("round-trips through unmergeRows back to a 1x1 region", () => {
    const t = table([[cell("a"), cell("b")], [cell("c"), cell("d")]]);
    const g = buildTableGrid(t);
    const rect = normalizeRect(g, { r0: 0, c0: 0, r1: 1, c1: 1 });
    const mergedRows = mergeRows(g, rect, merged("M", { col: 2, row: 2 }));
    const gm = buildTableGrid({ ...t, rows: mergedRows });
    let e = 0;
    const split = unmergeRows(gm, 0, 0, { id: "M", blocks: gm.slots[0]![0]!.cell.blocks }, () => cell(`e${e++}`));
    // Back to a full 2x2 grid of single cells.
    const gs = buildTableGrid({ ...t, rows: split });
    expect([gs.rows, gs.cols]).toEqual([2, 2]);
    expect(gs.slots[0]![0]!.cell.id).toBe("M");
    expect(gs.slots[1]![1]!.colSpan).toBe(1);
    expect(gs.slots[1]![1]!.rowSpan).toBe(1);
  });
});

describe("gridOriginOfCell", () => {
  it("maps array indices to the grid origin past a rowSpan hole", () => {
    const t = table([[cell("A", { row: 2 }), cell("B")], [cell("C")]]);
    const g = buildTableGrid(t);
    // C is rows[1].cells[0] but sits at grid col 1.
    expect(gridOriginOfCell(g, 1, 0)).toEqual({ row: 1, col: 1 });
  });
});
