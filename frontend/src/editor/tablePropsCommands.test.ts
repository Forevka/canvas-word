import { describe, expect, it } from "vitest";
import type { Document, SectionProps, TableBlock, TableCell } from "@cw/shared";
import { applyOp } from "@cw/shared";
import {
  insertTableColumnCmd,
  insertTableRowCmd,
  setCellTextDirectionCmd,
  setCellVAlignCmd,
  setRowHeightAtSelectionCmd,
  setRowPropsCmd,
  setTablePropsAtSelectionCmd,
} from "./commands";
import type { CellSelection, Command, EditorState } from "./state";

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
const tableOf = (doc: Document) => doc.blocks[0] as TableBlock;

const sel = (tableId: string, a: [number, number], f: [number, number]): CellSelection => ({
  tableId,
  anchor: { row: a[0], col: a[1] },
  focus: { row: f[0], col: f[1] },
});

/** Apply a command, returning the resulting doc AND a thunk that undoes it (by
 *  replaying the per-op inverses in reverse), so each test can assert round-trip. */
const applyWithUndo = (state: EditorState, cmd: Command): { doc: Document; undo: () => Document } => {
  const trn = cmd(state);
  if (!trn) throw new Error("command produced no transaction");
  let doc = state.doc;
  const inverses: Parameters<typeof applyOp>[1][] = [];
  for (const op of trn.ops) {
    const res = applyOp(doc, op);
    doc = res.doc;
    inverses.unshift(res.inverse);
  }
  return {
    doc,
    undo: () => {
      let d = doc;
      for (const inv of inverses) d = applyOp(d, inv).doc;
      return d;
    },
  };
};

const caretInCell = (table: TableBlock, ri: number, ci: number): EditorState => {
  const p = table.rows[ri]!.cells[ci]!.blocks[0]!;
  return {
    doc: docWith(table),
    selection: { anchor: { blockId: p.id, offset: 0 }, focus: { blockId: p.id, offset: 0 } },
    cellSelection: null,
  };
};

describe("setCellVAlignCmd", () => {
  it("sets vAlign on every cell in a range and undoes cleanly", () => {
    const table = grid2x2();
    const state: EditorState = { doc: docWith(table), selection: null, cellSelection: sel("t1", [0, 0], [1, 1]) };
    const { doc, undo } = applyWithUndo(state, setCellVAlignCmd("center"));
    expect(tableOf(doc).rows[0]!.cells[0]!.vAlign).toBe("center");
    expect(tableOf(doc).rows[1]!.cells[1]!.vAlign).toBe("center");
    const back = undo();
    expect(tableOf(back).rows[0]!.cells[0]!.vAlign).toBeUndefined();
    expect(tableOf(back).rows[1]!.cells[1]!.vAlign).toBeUndefined();
  });

  it('clears the field when set to the default "top"', () => {
    const table = grid2x2();
    const set = applyWithUndo(
      { doc: docWith(table), selection: null, cellSelection: sel("t1", [0, 0], [0, 0]) },
      setCellVAlignCmd("bottom"),
    ).doc;
    expect(tableOf(set).rows[0]!.cells[0]!.vAlign).toBe("bottom");
    const cleared = applyWithUndo(
      { doc: set, selection: null, cellSelection: sel("t1", [0, 0], [0, 0]) },
      setCellVAlignCmd("top"),
    ).doc;
    expect(tableOf(cleared).rows[0]!.cells[0]!.vAlign).toBeUndefined();
  });

  it("preserves a row's height props when editing a cell", () => {
    const table = grid2x2();
    table.rows[0]!.props = { height: { value: 60, rule: "atLeast" } };
    const { doc } = applyWithUndo(
      { doc: docWith(table), selection: null, cellSelection: sel("t1", [0, 0], [0, 1]) },
      setCellVAlignCmd("center"),
    );
    expect(tableOf(doc).rows[0]!.props?.height).toEqual({ value: 60, rule: "atLeast" });
  });
});

describe("setCellTextDirectionCmd", () => {
  it("sets and clears textDirection over the caret cell", () => {
    const table = grid2x2();
    const set = applyWithUndo(caretInCell(table, 0, 0), setCellTextDirectionCmd("tbRl")).doc;
    expect(tableOf(set).rows[0]!.cells[0]!.textDirection).toBe("tbRl");
    const cleared = applyWithUndo(caretInCell(tableOf(set), 0, 0), setCellTextDirectionCmd("lrTb")).doc;
    expect(tableOf(cleared).rows[0]!.cells[0]!.textDirection).toBeUndefined();
  });
});

describe("setRowPropsCmd", () => {
  it("toggles cantSplit / repeatHeader and undoes", () => {
    const table = grid2x2();
    const { doc, undo } = applyWithUndo(caretInCell(table, 0, 0), setRowPropsCmd({ repeatHeader: true }));
    expect(tableOf(doc).rows[0]!.props?.repeatHeader).toBe(true);
    expect(tableOf(undo()).rows[0]!.props).toBeUndefined();
  });

  it("clearing the only flag drops the props object", () => {
    const table = grid2x2();
    const set = applyWithUndo(caretInCell(table, 0, 0), setRowPropsCmd({ cantSplit: true })).doc;
    expect(tableOf(set).rows[0]!.props?.cantSplit).toBe(true);
    const cleared = applyWithUndo(caretInCell(tableOf(set), 0, 0), setRowPropsCmd({ cantSplit: false })).doc;
    expect(tableOf(cleared).rows[0]!.props).toBeUndefined();
  });

  it("keeps an existing height when toggling a flag", () => {
    const table = grid2x2();
    table.rows[0]!.props = { height: { value: 50, rule: "exact" } };
    const set = applyWithUndo(caretInCell(table, 0, 0), setRowPropsCmd({ cantSplit: true })).doc;
    expect(tableOf(set).rows[0]!.props).toEqual({ height: { value: 50, rule: "exact" }, cantSplit: true });
  });
});

describe("setRowHeightAtSelectionCmd", () => {
  it("sets the row height over a multi-row selection and undoes", () => {
    const table = grid2x2();
    const { doc, undo } = applyWithUndo(
      { doc: docWith(table), selection: null, cellSelection: sel("t1", [0, 0], [1, 0]) },
      setRowHeightAtSelectionCmd({ value: 80, rule: "atLeast" }),
    );
    expect(tableOf(doc).rows[0]!.props?.height).toEqual({ value: 80, rule: "atLeast" });
    expect(tableOf(doc).rows[1]!.props?.height).toEqual({ value: 80, rule: "atLeast" });
    expect(tableOf(undo()).rows[0]!.props?.height).toBeUndefined();
  });

  it("clears the height with null", () => {
    const table = grid2x2();
    table.rows[0]!.props = { height: { value: 80, rule: "atLeast" } };
    const cleared = applyWithUndo(caretInCell(table, 0, 0), setRowHeightAtSelectionCmd(null)).doc;
    expect(tableOf(cleared).rows[0]!.props?.height).toBeUndefined();
  });
});

describe("setTablePropsAtSelectionCmd", () => {
  it("sets table indent and undoes", () => {
    const table = grid2x2();
    const { doc, undo } = applyWithUndo(caretInCell(table, 0, 0), setTablePropsAtSelectionCmd({ indentPx: 48 }));
    expect(tableOf(doc).indentPx).toBe(48);
    expect(tableOf(undo()).indentPx).toBeUndefined();
  });

  it("sets and clears the table-level default shading / borders / cell margin", () => {
    const table = grid2x2();
    const set = applyWithUndo(
      caretInCell(table, 0, 0),
      setTablePropsAtSelectionCmd({
        defaultShading: "#eef",
        defaultBorders: { top: { color: "#123456", widthPx: 1 } },
        defaultCellMargin: { top: 2, right: 7, bottom: 2, left: 7 },
      }),
    ).doc;
    expect(tableOf(set).defaultShading).toBe("#eef");
    expect(tableOf(set).defaultBorders?.top?.color).toBe("#123456");
    expect(tableOf(set).defaultCellMargin).toEqual({ top: 2, right: 7, bottom: 2, left: 7 });
    const cleared = applyWithUndo(
      caretInCell(tableOf(set), 0, 0),
      setTablePropsAtSelectionCmd({ defaultShading: null, defaultBorders: null, defaultCellMargin: null }),
    ).doc;
    expect(cleared.blocks[0]).toMatchObject({ kind: "table" });
    expect((cleared.blocks[0] as TableBlock).defaultShading).toBeUndefined();
    expect((cleared.blocks[0] as TableBlock).defaultBorders).toBeUndefined();
    expect((cleared.blocks[0] as TableBlock).defaultCellMargin).toBeUndefined();
  });
});

describe("insert row/column inherits the proto cell's format", () => {
  // Table-level defaults are baked onto every cell at import/build time, so a new
  // cell must carry over the neighbour's borders/shading/margin/vAlign or it falls
  // back to the engine's bare defaults (light grid, no fill, default padding).
  const fmtCell = (text: string): TableCell => ({
    ...cell(text),
    shading: "#ffeeee",
    borders: { top: { color: "#ff0000", widthPx: 1 }, bottom: { color: "#ff0000", widthPx: 1 } } as never,
    margin: { top: 2, right: 8, bottom: 2, left: 8 } as never,
    vAlign: "center",
  });

  it("Insert Row Below copies borders/shading/margin/vAlign into the new row", () => {
    const table: TableBlock = { kind: "table", id: "t1", revision: 0, rows: [{ cells: [fmtCell("a"), fmtCell("b")] }] };
    const { doc, undo } = applyWithUndo(caretInCell(table, 0, 0), insertTableRowCmd("below"));
    const t = tableOf(doc);
    expect(t.rows.length).toBe(2);
    for (const c of t.rows[1]!.cells) {
      expect(c.shading).toBe("#ffeeee");
      expect(c.borders).toBeDefined();
      expect(c.margin).toEqual({ top: 2, right: 8, bottom: 2, left: 8 });
      expect(c.vAlign).toBe("center");
    }
    // Deep-cloned — the new row must not share the proto's border/margin objects.
    expect(t.rows[1]!.cells[0]!.borders).not.toBe(t.rows[0]!.cells[0]!.borders);
    expect(t.rows[1]!.cells[0]!.margin).not.toBe(t.rows[0]!.cells[0]!.margin);
    expect(tableOf(undo()).rows.length).toBe(1);
  });

  it("Insert Column Right copies the proto column's format into every new cell", () => {
    const table: TableBlock = {
      kind: "table", id: "t1", revision: 0,
      rows: [{ cells: [fmtCell("a")] }, { cells: [fmtCell("c")] }],
    };
    const { doc } = applyWithUndo(caretInCell(table, 0, 0), insertTableColumnCmd("right"));
    const t = tableOf(doc);
    expect(t.rows.every((r) => r.cells.length === 2)).toBe(true);
    for (const r of t.rows) {
      const added = r.cells[1]!;
      expect(added.shading).toBe("#ffeeee");
      expect(added.margin).toEqual({ top: 2, right: 8, bottom: 2, left: 8 });
      expect(added.vAlign).toBe("center");
      expect(added.borders).toBeDefined();
    }
  });

  it("a plain cell with no format yields a plain new cell (no spurious props)", () => {
    const { doc } = applyWithUndo(caretInCell(grid2x2(), 0, 0), insertTableRowCmd("below"));
    const added = tableOf(doc).rows[1]!.cells[0]!;
    expect(added.shading).toBeUndefined();
    expect(added.borders).toBeUndefined();
    expect(added.margin).toBeUndefined();
  });
});
