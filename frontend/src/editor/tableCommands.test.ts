import { describe, expect, it } from "vitest";
import type { Document, SectionProps, TableBlock, TableCell } from "@cw/shared";
import { applyOp, buildTableGrid } from "@cw/shared";
import { mergeCellsCmd, setCellsBordersCmd, setCellsShadingCmd, toggleList, unmergeCellCmd } from "./commands";
import type { CellBorder } from "@cw/shared";
import type { CellSelection, EditorState } from "./state";

let n = 0;
const cell = (text: string): TableCell => ({
  id: `c${n++}`,
  blocks: [{ kind: "paragraph", id: `p${n++}`, revision: 0, runs: [{ text, style: char() }], style: para() }],
});
const char = () =>
  ({ fontFamily: "Georgia, serif", fontSizePx: 14, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" }) as never;
const para = () => ({ align: "left", lineHeight: 1.35, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 }) as never;

const SECTION: SectionProps = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
};

const grid2x2 = (): TableBlock => ({
  kind: "table",
  id: "t1",
  revision: 0,
  rows: [
    { cells: [cell("a"), cell("b")] },
    { cells: [cell("c"), cell("d")] },
  ],
});

const docWith = (table: TableBlock): Document => ({ section: SECTION, blocks: [table] });

const runCmd = (state: EditorState, cmd: ReturnType<typeof mergeCellsCmd>): Document => {
  const trn = cmd(state);
  if (!trn) throw new Error("command produced no transaction");
  let doc = state.doc;
  for (const op of trn.ops) doc = applyOp(doc, op).doc;
  return doc;
};

const sel = (tableId: string, a: [number, number], f: [number, number]): CellSelection => ({
  tableId,
  anchor: { row: a[0], col: a[1] },
  focus: { row: f[0], col: f[1] },
});

describe("mergeCellsCmd (rectangular)", () => {
  it("merges a full 2x2 block into one spanning cell carrying all content", () => {
    const table = grid2x2();
    const state: EditorState = { doc: docWith(table), selection: null, cellSelection: sel("t1", [0, 0], [1, 1]) };
    const doc = runCmd(state, mergeCellsCmd());
    const g = buildTableGrid(doc.blocks[0] as TableBlock);
    // One cell owns the whole grid now.
    const owner = g.slots[0]![0]!.cell;
    expect([owner.colSpan, owner.rowSpan]).toEqual([2, 2]);
    expect(g.slots[1]![1]!.cell).toBe(owner);
    const texts = owner.blocks.map((b) => (b.kind === "paragraph" ? b.runs.map((r) => r.text).join("") : ""));
    expect(texts).toEqual(["a", "b", "c", "d"]);
  });

  it("merges a vertical pair (rowSpan) and empties the covered row's column", () => {
    const table = grid2x2();
    const state: EditorState = { doc: docWith(table), selection: null, cellSelection: sel("t1", [0, 0], [1, 0]) };
    const doc = runCmd(state, mergeCellsCmd());
    const t = doc.blocks[0] as TableBlock;
    const g = buildTableGrid(t);
    expect(g.slots[1]![0]!.cell).toBe(g.slots[0]![0]!.cell); // col 0 continues down
    expect(g.slots[0]![0]!.rowSpan).toBe(2);
    expect(g.slots[0]![0]!.colSpan ?? 1).toBe(1);
    // row 1 now has just its column-1 cell.
    expect(t.rows[1]!.cells.length).toBe(1);
  });

  it("is a no-op for a single-cell selection", () => {
    const table = grid2x2();
    const state: EditorState = { doc: docWith(table), selection: null, cellSelection: sel("t1", [0, 0], [0, 0]) };
    expect(mergeCellsCmd()(state)).toBeNull();
  });
});

const BLUE: CellBorder = { color: "#0000ff", widthPx: 1 };
const grid3x3 = (): TableBlock => ({
  kind: "table",
  id: "t1",
  revision: 0,
  rows: [
    { cells: [cell("a"), cell("b"), cell("c")] },
    { cells: [cell("d"), cell("e"), cell("f")] },
    { cells: [cell("g"), cell("h"), cell("i")] },
  ],
});
const tableOf = (doc: Document) => doc.blocks[0] as TableBlock;

describe("setCellsBordersCmd edge presets", () => {
  it("Outside applies only to the perimeter edges of the 2x2 selection", () => {
    const table = grid3x3();
    const edges = { top: true, right: true, bottom: true, left: true }; // Outside
    const doc = runCmd(
      { doc: docWith(table), selection: null, cellSelection: sel("t1", [0, 0], [1, 1]) },
      setCellsBordersCmd(BLUE, edges),
    );
    const rows = tableOf(doc).rows;
    const tl = rows[0]!.cells[0]!; // (0,0): outer top + outer left, NOT right/bottom (interior)
    expect(tl.borders?.top?.color).toBe("#0000ff");
    expect(tl.borders?.left?.color).toBe("#0000ff");
    expect(tl.borders?.right).toBeUndefined();
    expect(tl.borders?.bottom).toBeUndefined();
    const br = rows[1]!.cells[1]!; // (1,1): outer right + outer bottom only
    expect(br.borders?.right?.color).toBe("#0000ff");
    expect(br.borders?.bottom?.color).toBe("#0000ff");
    expect(br.borders?.top).toBeUndefined();
    expect(br.borders?.left).toBeUndefined();
    // A cell outside the selection is untouched.
    expect(rows[2]!.cells[2]!.borders).toBeUndefined();
  });

  it("Inside applies only to the shared interior edges", () => {
    const table = grid3x3();
    const doc = runCmd(
      { doc: docWith(table), selection: null, cellSelection: sel("t1", [0, 0], [1, 1]) },
      setCellsBordersCmd(BLUE, { insideH: true, insideV: true }),
    );
    const rows = tableOf(doc).rows;
    const tl = rows[0]!.cells[0]!; // interior edges of (0,0) are its right + bottom
    expect(tl.borders?.right?.color).toBe("#0000ff");
    expect(tl.borders?.bottom?.color).toBe("#0000ff");
    expect(tl.borders?.top).toBeUndefined(); // perimeter — untouched by Inside
    expect(tl.borders?.left).toBeUndefined();
  });

  it("None (spec=null, all flags) clears edges but keeps an explicit borders object", () => {
    const table = grid3x3();
    const all = { top: true, right: true, bottom: true, left: true, insideH: true, insideV: true };
    const bordered = runCmd(
      { doc: docWith(table), selection: null, cellSelection: sel("t1", [0, 0], [1, 1]) },
      setCellsBordersCmd(BLUE, all),
    );
    const cleared = runCmd(
      { doc: bordered, selection: null, cellSelection: sel("t1", [0, 0], [1, 1]) },
      setCellsBordersCmd(null, all),
    );
    const tl = tableOf(cleared).rows[0]!.cells[0]!;
    expect(tl.borders).toBeDefined(); // present — suppresses the gray default grid
    expect(tl.borders?.top).toBeUndefined();
    expect(tl.borders?.right).toBeUndefined();
  });
});

describe("setCellsShadingCmd", () => {
  it("fills every cell in the range and clears with null", () => {
    const table = grid3x3();
    const filled = runCmd(
      { doc: docWith(table), selection: null, cellSelection: sel("t1", [0, 0], [1, 1]) },
      setCellsShadingCmd("#ffff00"),
    );
    expect(tableOf(filled).rows[0]!.cells[0]!.shading).toBe("#ffff00");
    expect(tableOf(filled).rows[1]!.cells[1]!.shading).toBe("#ffff00");
    expect(tableOf(filled).rows[2]!.cells[2]!.shading).toBeUndefined(); // outside the range
    const cleared = runCmd(
      { doc: filled, selection: null, cellSelection: sel("t1", [0, 0], [0, 0]) },
      setCellsShadingCmd(null),
    );
    expect(tableOf(cleared).rows[0]!.cells[0]!.shading).toBeUndefined();
  });
});

describe("lists in table cells", () => {
  it("toggleList applies a list to a paragraph inside a cell", () => {
    const table = grid2x2();
    const cellPara = table.rows[0]!.cells[0]!.blocks[0]!;
    const id = cellPara.kind === "paragraph" ? cellPara.id : "";
    const state: EditorState = {
      doc: docWith(table),
      selection: { anchor: { blockId: id, offset: 0 }, focus: { blockId: id, offset: 0 } },
      cellSelection: null,
    };
    const doc = runCmd(state, toggleList("bullet"));
    const para = (doc.blocks[0] as TableBlock).rows[0]!.cells[0]!.blocks[0]!;
    expect(para.kind === "paragraph" && para.style.list?.listId).toBe("bullets");
    expect(doc.lists?.["bullets"]).toBeDefined(); // definition materialized
  });
});

describe("unmergeCellCmd (rectangular)", () => {
  it("round-trips a 2x2 merge back to four single cells", () => {
    const table = grid2x2();
    const merged = runCmd(
      { doc: docWith(table), selection: null, cellSelection: sel("t1", [0, 0], [1, 1]) },
      mergeCellsCmd(),
    );
    const split = runCmd(
      { doc: merged, selection: null, cellSelection: sel("t1", [0, 0], [0, 0]) },
      unmergeCellCmd(),
    );
    const g = buildTableGrid(split.blocks[0] as TableBlock);
    expect([g.rows, g.cols]).toEqual([2, 2]);
    for (let r = 0; r < 2; r++)
      for (let c = 0; c < 2; c++) {
        expect(g.slots[r]![c]!.colSpan).toBe(1);
        expect(g.slots[r]![c]!.rowSpan).toBe(1);
      }
    // Content stayed in the top-left.
    const tl = g.slots[0]![0]!.cell.blocks.map((b) => (b.kind === "paragraph" ? b.runs.map((r) => r.text).join("") : ""));
    expect(tl).toEqual(["a", "b", "c", "d"]);
  });
});
