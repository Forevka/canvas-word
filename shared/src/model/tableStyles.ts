// Table styles — a round-trippable mirror of OOXML w:style[@w:type="table"] with
// conditional formatting (w:tblStylePr): a whole-table base plus per-band overrides
// (header/total row, first/last column, row/column banding, corner cells). Like
// named styles, the editor MODEL stays concrete: applying or editing a table style
// BAKES the effective per-cell props onto TableCell.shading/borders/margin (the
// layout engine reads those directly), while the TableStyle entity + TableBlock.styleId
// are retained for re-editing and docx round-trip.

import type { CellBorders, CellMargin, CharStyle, ParaStyle, TableBlock, TableCell, TableCondOverrides, TableRow } from "./document";
import { buildTableGrid, rebuildRows } from "./tableGrid";

/** Conditional-format slots (OOXML w:tblStylePr w:type values). "wholeTable" is
 *  the base applied to every cell. */
export type TableCond =
  | "wholeTable"
  | "firstRow" | "lastRow" | "firstCol" | "lastCol"
  | "band1Horz" | "band2Horz" | "band1Vert" | "band2Vert"
  | "nwCell" | "neCell" | "swCell" | "seCell";

/** Formatting for one conditional band. Borders/shading/margin map to cell
 *  props; char/para map to the cell's content runs/paragraphs. */
export interface TableCondProps {
  char?: Partial<CharStyle>;
  para?: Partial<ParaStyle>;
  shading?: string;
  borders?: CellBorders;
  margin?: CellMargin;
}

export interface TableStyle {
  id: string; // shares the docx styleId space
  name: string;
  basedOn?: string;
  /** Per-condition formatting; `wholeTable` is the base. */
  conds: Partial<Record<TableCond, TableCondProps>>;
  /** Banding stripe sizes (rows/cols per stripe). Default 1. */
  rowBandSize?: number;
  colBandSize?: number;
}

/** Which conditional bands a given cell participates in, derived from its grid
 *  position + the table's tblLook flags. */
export interface CellCondFlags {
  firstRow?: boolean;
  lastRow?: boolean;
  firstCol?: boolean;
  lastCol?: boolean;
  /** Row band parity for banded rows (1 = band1Horz, 2 = band2Horz). */
  rowBand?: 1 | 2;
  /** Column band parity for banded cols. */
  colBand?: 1 | 2;
}

export function tableStyleById(styles: Record<string, TableStyle>, id: string): TableStyle | undefined {
  return styles[id];
}

/** Collapse a table style's basedOn chain into one set of per-condition props
 *  (root → leaf, child overriding per condition/field). Cycle-guarded. */
export function resolveTableStyle(styles: Record<string, TableStyle>, id: string): Partial<Record<TableCond, TableCondProps>> {
  const chain: TableStyle[] = [];
  const seen = new Set<string>();
  for (let cur = styles[id]; cur && !seen.has(cur.id); cur = cur.basedOn ? styles[cur.basedOn] : undefined) {
    seen.add(cur.id);
    chain.unshift(cur);
  }
  const out: Partial<Record<TableCond, TableCondProps>> = {};
  for (const s of chain) {
    for (const [cond, props] of Object.entries(s.conds) as [TableCond, TableCondProps][]) {
      const prev = out[cond] ?? {};
      out[cond] = mergeCond(prev, props);
    }
  }
  return out;
}

function mergeCond(base: TableCondProps, add: TableCondProps): TableCondProps {
  const out: TableCondProps = { ...base };
  if (add.char) out.char = { ...base.char, ...add.char };
  if (add.para) out.para = { ...base.para, ...add.para };
  if (add.shading !== undefined) out.shading = add.shading;
  if (add.borders) out.borders = { ...base.borders, ...add.borders };
  if (add.margin) out.margin = add.margin;
  return out;
}

/** OOXML conditional-formatting precedence (least → most specific), so later
 *  entries override earlier when combining for a cell. Corners are most specific. */
const COND_ORDER: TableCond[] = [
  "wholeTable",
  "band1Horz", "band2Horz", "band1Vert", "band2Vert",
  "firstRow", "lastRow", "firstCol", "lastCol",
  "nwCell", "neCell", "swCell", "seCell",
];

/** The effective cell props for one cell: layer the table style's applicable
 *  conditions in OOXML precedence order. `direct` (the cell's own formatting)
 *  always wins last. Returns only the cell-level fields (shading/borders/margin
 *  + the cell's content char/para deltas). */
export function effectiveCellProps(
  resolved: Partial<Record<TableCond, TableCondProps>>,
  flags: CellCondFlags,
  direct?: { shading?: string; borders?: CellBorders; margin?: CellMargin },
): TableCondProps {
  const active = (cond: TableCond): boolean => {
    switch (cond) {
      case "wholeTable": return true;
      case "firstRow": return !!flags.firstRow;
      case "lastRow": return !!flags.lastRow;
      case "firstCol": return !!flags.firstCol;
      case "lastCol": return !!flags.lastCol;
      case "band1Horz": return flags.rowBand === 1;
      case "band2Horz": return flags.rowBand === 2;
      case "band1Vert": return flags.colBand === 1;
      case "band2Vert": return flags.colBand === 2;
      case "nwCell": return !!flags.firstRow && !!flags.firstCol;
      case "neCell": return !!flags.firstRow && !!flags.lastCol;
      case "swCell": return !!flags.lastRow && !!flags.firstCol;
      case "seCell": return !!flags.lastRow && !!flags.lastCol;
    }
  };
  let out: TableCondProps = {};
  for (const cond of COND_ORDER) {
    if (active(cond) && resolved[cond]) out = mergeCond(out, resolved[cond]!);
  }
  if (direct) {
    if (direct.shading !== undefined) out.shading = direct.shading;
    if (direct.borders) out.borders = { ...out.borders, ...direct.borders };
    if (direct.margin) out.margin = direct.margin;
  }
  return out;
}

/** Compute a cell's conditional flags from grid position + the table's band
 *  toggles (w:tblLook). rowCount/colCount are the grid dimensions. */
export function cellCondFlags(
  row: number, col: number, rowCount: number, colCount: number,
  opts: { firstRow?: boolean; lastRow?: boolean; firstCol?: boolean; lastCol?: boolean; bandRows?: boolean; bandCols?: boolean; rowBandSize?: number; colBandSize?: number },
): CellCondFlags {
  const flags: CellCondFlags = {};
  const isFirstRow = opts.firstRow && row === 0;
  const isLastRow = opts.lastRow && row === rowCount - 1;
  const isFirstCol = opts.firstCol && col === 0;
  const isLastCol = opts.lastCol && col === colCount - 1;
  if (isFirstRow) flags.firstRow = true;
  if (isLastRow) flags.lastRow = true;
  if (isFirstCol) flags.firstCol = true;
  if (isLastCol) flags.lastCol = true;
  // Banding applies to the "body" rows/cols (excluding header/total/first/last).
  if (opts.bandRows && !isFirstRow && !isLastRow) {
    const headerOffset = opts.firstRow ? 1 : 0;
    const band = Math.floor((row - headerOffset) / Math.max(1, opts.rowBandSize ?? 1));
    flags.rowBand = band % 2 === 0 ? 1 : 2;
  }
  if (opts.bandCols && !isFirstCol && !isLastCol) {
    const colOffset = opts.firstCol ? 1 : 0;
    const band = Math.floor((col - colOffset) / Math.max(1, opts.colBandSize ?? 1));
    flags.colBand = band % 2 === 0 ? 1 : 2;
  }
  return flags;
}

/** Word's default w:tblLook: header row on, row banding on. */
export const DEFAULT_TBL_LOOK: TableCondOverrides = { firstRow: true, bandRows: true };

/** Apply a band's props onto a concrete style, but only where a property was not
 *  explicitly set — so explicit formatting keeps precedence while default-valued
 *  props pick up the band. "Explicitly set" is determined first by `explicit` (an
 *  author-supplied-key set carried from the builder as provenance) and then, as a
 *  fallback for fields with no provenance, by value-equality with `defaults`. The
 *  provenance set is what lets an explicit value that HAPPENS to equal the default
 *  survive a conflicting band (the value-equality heuristic alone can't tell that
 *  case apart from genuinely-unset content). */
function applyBandProps<T extends object>(current: T, patch: Partial<T>, defaults: T, explicit?: ReadonlySet<string>): T {
  const cur = current as Record<string, unknown>;
  const def = defaults as Record<string, unknown>;
  const out: Record<string, unknown> = { ...cur };
  for (const [k, v] of Object.entries(patch as Record<string, unknown>)) {
    if (explicit?.has(k)) continue; // author explicitly set this field — keep it, even if it equals the default
    if (v !== undefined && (cur[k] === undefined || cur[k] === def[k])) out[k] = v;
  }
  return out as T;
}

/** Recompute every cell's baked shading/borders/margin from a resolved table style
 *  + the table's active bands (span-aware). The TableStyle entity + TableBlock.styleId
 *  are kept for re-editing and docx round-trip; the layout engine reads the baked
 *  concrete cell props. Applying a style replaces direct cell shading/borders/margin
 *  (like Word).
 *
 *  When `contentDefaults` is given, the band's char/para are ALSO materialized as
 *  concrete run/paragraph styles on the cell content — but only where a property
 *  was not explicitly set (see `applyBandProps`), so explicit content formatting
 *  keeps precedence. Omit it to leave cell content untouched.
 *
 *  `explicitChar`/`explicitPara` carry per-object provenance from the builder: a
 *  map from a content run / paragraph block to the set of style keys the author
 *  supplied. Object identity is preserved through the grid rebuild, so the bake
 *  can look each run/paragraph up. This is what preserves an explicit value that
 *  equals the resolved default against a conflicting band (issue #30); without
 *  provenance the bake falls back to value-equality only. */
export function bakeTableStyleRows(
  table: TableBlock,
  style: TableStyle,
  styles: Record<string, TableStyle>,
  overrides: TableCondOverrides,
  contentDefaults?: {
    char: CharStyle;
    para: ParaStyle;
    explicitChar?: WeakMap<object, ReadonlySet<string>>;
    explicitPara?: WeakMap<object, ReadonlySet<string>>;
  },
): TableRow[] {
  const resolved = resolveTableStyle(styles, style.id);
  const grid = buildTableGrid(table);
  return rebuildRows(grid, (cell, s) => {
    const flags = cellCondFlags(s.originRow, s.originCol, grid.rows, grid.cols, {
      ...overrides,
      rowBandSize: style.rowBandSize ?? 1,
      colBandSize: style.colBandSize ?? 1,
    });
    const props = effectiveCellProps(resolved, flags);
    const next: TableCell = { ...cell };
    if (props.shading !== undefined) next.shading = props.shading;
    else delete next.shading;
    if (props.borders && Object.keys(props.borders).length > 0) next.borders = props.borders;
    else delete next.borders;
    if (props.margin) next.margin = props.margin;
    else delete next.margin;
    if (contentDefaults && (props.char || props.para)) {
      next.blocks = next.blocks.map((b) =>
        b.kind === "paragraph"
          ? {
              ...b,
              style: props.para ? applyBandProps(b.style, props.para, contentDefaults.para, contentDefaults.explicitPara?.get(b)) : b.style,
              runs: props.char
                ? b.runs.map((r) => ({ ...r, style: applyBandProps(r.style, props.char!, contentDefaults.char, contentDefaults.explicitChar?.get(r)) }))
                : b.runs,
            }
          : b,
      );
    }
    return next;
  });
}

export const TABLE_COND_LABELS: Record<TableCond, string> = {
  wholeTable: "Whole table",
  firstRow: "Header row",
  lastRow: "Total row",
  firstCol: "First column",
  lastCol: "Last column",
  band1Horz: "Banded rows (odd)",
  band2Horz: "Banded rows (even)",
  band1Vert: "Banded columns (odd)",
  band2Vert: "Banded columns (even)",
  nwCell: "Top-left cell",
  neCell: "Top-right cell",
  swCell: "Bottom-left cell",
  seCell: "Bottom-right cell",
};
