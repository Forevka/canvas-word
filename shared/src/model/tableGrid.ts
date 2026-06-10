// A span-aware GRID view of a TableBlock. The model stores cells HTML-style: a
// rowSpan owner lives in its TOP row and the rows it spans simply omit a cell for
// those columns. That's compact but awkward for anything that reasons in 2-D
// (rectangular selection, merge, edge-aware borders). buildTableGrid materializes
// the full grid — every grid coordinate points at the cell that occupies it —
// while remembering each cell's array indices so callers can rebuild the rows.

import type { TableBlock, TableCell, TableRow } from "./document";
import { gridColumnCount } from "./ops";

/** One grid coordinate's occupant. Several adjacent slots can reference the same
 *  cell (a merged cell fills its whole rectangle). */
export interface GridSlot {
  cell: TableCell;
  /** Array indices of the OWNING cell: table.rows[ri].cells[ci]. */
  ri: number;
  ci: number;
  /** Grid origin (top-left) of the owning cell. */
  originRow: number;
  originCol: number;
  rowSpan: number;
  colSpan: number;
}

export interface TableGrid {
  rows: number;
  cols: number;
  /** slots[gridRow][gridCol] — undefined only for a malformed (ragged) table. */
  slots: (GridSlot | undefined)[][];
}

/** Materialize the full grid. Grid rows map 1:1 to table.rows; grid columns come
 *  from gridColumnCount (the widest row measured in grid columns). */
export function buildTableGrid(t: TableBlock): TableGrid {
  const cols = gridColumnCount(t);
  const rows = t.rows.length;
  const slots: (GridSlot | undefined)[][] = Array.from({ length: rows }, () =>
    new Array<GridSlot | undefined>(cols).fill(undefined),
  );
  for (let ri = 0; ri < rows; ri++) {
    let col = 0;
    const cells = t.rows[ri]!.cells;
    for (let ci = 0; ci < cells.length; ci++) {
      const cell = cells[ci]!;
      // Skip columns already claimed by a rowSpan from an earlier row.
      while (col < cols && slots[ri]![col] !== undefined) col++;
      if (col >= cols) break; // ragged overflow — drop the rest
      const colSpan = Math.max(1, cell.colSpan ?? 1);
      const rowSpan = Math.max(1, cell.rowSpan ?? 1);
      const slot: GridSlot = { cell, ri, ci, originRow: ri, originCol: col, rowSpan, colSpan };
      for (let r = 0; r < rowSpan && ri + r < rows; r++) {
        for (let c = 0; c < colSpan && col + c < cols; c++) {
          slots[ri + r]![col + c] = slot;
        }
      }
      col += colSpan;
    }
  }
  return { rows, cols, slots };
}

/** A rectangular block of grid coordinates (inclusive on both ends). */
export interface GridRect {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

/** Grow a grid rectangle until it fully contains every cell it touches: a merged
 *  cell poking past an edge pulls that edge out to the cell's full extent. Word
 *  does the same — you can't half-select a merged cell. Iterates to a fixed
 *  point (one expansion can expose another straddling cell). */
export function normalizeRect(grid: TableGrid, rect: GridRect): GridRect {
  let { r0, c0, r1, c1 } = rect;
  r0 = clamp(r0, 0, grid.rows - 1);
  r1 = clamp(r1, 0, grid.rows - 1);
  c0 = clamp(c0, 0, grid.cols - 1);
  c1 = clamp(c1, 0, grid.cols - 1);
  if (r1 < r0) [r0, r1] = [r1, r0];
  if (c1 < c0) [c0, c1] = [c1, c0];
  for (;;) {
    let grew = false;
    for (let r = r0; r <= r1; r++) {
      for (let c = c0; c <= c1; c++) {
        const s = grid.slots[r]![c];
        if (!s) continue;
        if (s.originRow < r0) (r0 = s.originRow), (grew = true);
        if (s.originCol < c0) (c0 = s.originCol), (grew = true);
        if (s.originRow + s.rowSpan - 1 > r1) (r1 = s.originRow + s.rowSpan - 1), (grew = true);
        if (s.originCol + s.colSpan - 1 > c1) (c1 = s.originCol + s.colSpan - 1), (grew = true);
      }
    }
    if (!grew) break;
  }
  return { r0, c0, r1, c1 };
}

/** The distinct owning cells inside a rectangle, in reading order (row-major by
 *  origin). Each cell appears once even when it spans many slots. */
export function cellsInRect(grid: TableGrid, rect: GridRect): GridSlot[] {
  const seen = new Set<TableCell>();
  const out: GridSlot[] = [];
  for (let r = rect.r0; r <= rect.r1; r++) {
    for (let c = rect.c0; c <= rect.c1; c++) {
      const s = grid.slots[r]![c];
      if (!s || seen.has(s.cell)) continue;
      seen.add(s.cell);
      out.push(s);
    }
  }
  return out;
}

/** Grid origin (top-left) of the cell at array indices table.rows[ri].cells[ci]. */
export function gridOriginOfCell(grid: TableGrid, ri: number, ci: number): { row: number; col: number } | null {
  for (let r = 0; r < grid.rows; r++) {
    for (let c = 0; c < grid.cols; c++) {
      const s = grid.slots[r]![c];
      if (s && s.ri === ri && s.ci === ci && s.originRow === r && s.originCol === c) return { row: r, col: c };
    }
  }
  return null;
}

/** Rebuild every row with the rectangle collapsed into `merged` (placed at the
 *  top-left). Cells fully inside the rectangle are absorbed; everything else
 *  keeps its place. Rows wholly inside the vertical span may come out empty —
 *  that's the valid HTML/OOXML "all columns continue from above" shape. The
 *  rectangle must already be normalized (no cell straddles an edge). */
export function mergeRows(grid: TableGrid, rect: GridRect, merged: TableCell): TableRow[] {
  const { r0, c0, r1, c1 } = rect;
  const out: TableRow[] = [];
  for (let r = 0; r < grid.rows; r++) {
    const cells: TableCell[] = [];
    for (let c = 0; c < grid.cols; c++) {
      const s = grid.slots[r]![c];
      if (!s || s.originRow !== r || s.originCol !== c) continue; // emit each cell once, at its origin
      const inRect = s.originRow >= r0 && s.originRow <= r1 && s.originCol >= c0 && s.originCol <= c1;
      if (inRect) {
        if (r === r0 && c === c0) cells.push(merged);
      } else {
        cells.push(s.cell);
      }
    }
    out.push({ cells });
  }
  return out;
}

/** Rebuild every row with the spanned cell at (originRow, originCol) split back
 *  into a 1x1 grid: `topLeft` keeps the content, `makeEmpty()` fills the rest. */
export function unmergeRows(
  grid: TableGrid,
  originRow: number,
  originCol: number,
  topLeft: TableCell,
  makeEmpty: () => TableCell,
): TableRow[] {
  const target = grid.slots[originRow]?.[originCol]?.cell;
  const out: TableRow[] = [];
  for (let r = 0; r < grid.rows; r++) {
    const cells: TableCell[] = [];
    for (let c = 0; c < grid.cols; c++) {
      const s = grid.slots[r]![c];
      if (!s) continue;
      if (s.cell === target) {
        cells.push(r === originRow && c === originCol ? topLeft : makeEmpty());
      } else if (s.originRow === r && s.originCol === c) {
        cells.push(s.cell);
      }
    }
    out.push({ cells });
  }
  return out;
}

/** Rebuild every row, passing each cell (once, at its grid origin) through
 *  `mapCell`. Structure is preserved exactly — use for per-cell property edits
 *  (borders, shading) that don't change the grid. */
export function rebuildRows(grid: TableGrid, mapCell: (cell: TableCell, slot: GridSlot) => TableCell): TableRow[] {
  const out: TableRow[] = [];
  for (let r = 0; r < grid.rows; r++) {
    const cells: TableCell[] = [];
    for (let c = 0; c < grid.cols; c++) {
      const s = grid.slots[r]![c];
      if (!s || s.originRow !== r || s.originCol !== c) continue;
      cells.push(mapCell(s.cell, s));
    }
    out.push({ cells });
  }
  return out;
}

const clamp = (v: number, lo: number, hi: number): number => Math.max(lo, Math.min(hi, v));
