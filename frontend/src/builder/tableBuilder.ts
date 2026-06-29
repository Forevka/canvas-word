// Table composition. Two entry shapes share this implementation:
//   - data-driven: .table([["Name", "DOB"], ["Ada", "1815"]], { headerRow: true })
//   - structural:  .table(t => t.row(r => r.cell("x", { colSpan: 2 })))
// Cells hold full block stories; a cell callback gets a StoryBuilder, so
// paragraphs/images/lists inside cells reuse the normal scope surface.

import type { Block, CellBorders, CellMargin, CharStyle, ParaStyle, RowProps, TableBlock, TableBorders, TableCell, TableCondOverrides, TableRow } from "@cw/shared";
import { bakeTableStyleRows, DEFAULT_TBL_LOOK } from "@cw/shared";
import type { BuilderContext } from "./blockFactory";
import { StoryBuilder } from "./storyBuilder";
import type { TableStylePreset } from "./tableStyles";

/** A cell in the data-driven shape: plain text, or text + cell properties. */
export interface CellSpec {
  text?: string;
  /** Columns this cell covers (HTML colspan semantics). */
  colSpan?: number;
  /** Rows this cell covers — spanned-into rows simply omit the cell. */
  rowSpan?: number;
  /** Background fill (CSS color). */
  shading?: string;
  /** Per-edge borders; absent = the renderer's default light grid. */
  borders?: CellBorders;
  /** Inner padding override, px per side. */
  margin?: CellMargin;
  /** Char formatting for the cell's text. */
  style?: Partial<CharStyle>;
  align?: ParaStyle["align"];
  /** Preferred cell width (w:tcW): absolute px or a percent of table width. */
  preferredWidth?: { px: number; type: "abs" | "pct" };
  /** Vertical alignment of the cell's content (w:vAlign). Absent = "top". Visible
   *  when the cell is taller than its content (a rowSpan or a tall sibling row). */
  vAlign?: "top" | "center" | "bottom";
  /** Text flow direction inside the cell (w:textDirection). Absent = "lrTb"
   *  (horizontal). "tbRl"/"btLr" are Word's rotated (vertical) cell text. */
  textDirection?: "lrTb" | "tbRl" | "btLr" | "lrTbV" | "tbRlV" | "tbLrV";
  /** Suppress content wrapping (w:noWrap). Absent = wrap. */
  noWrap?: boolean;
  /** Fit text to the cell width by tracking (w:tcFitText). Absent = off. */
  fitText?: boolean;
  /** Ignore the end-of-cell mark for row-height (w:hideMark). Absent = off. */
  hideMark?: boolean;
}

export type CellContent = string | CellSpec;

export type CellOptions = Omit<CellSpec, "text">;

/** Row-level properties (w:trPr) for a single .row(). */
export interface RowOptions {
  /** Fixed/minimum row height in px (w:trHeight). Pair with `heightRule`. */
  height?: number;
  /** How `height` is enforced: "atLeast" (min, grows with content — default) or
   *  "exact" (pinned; taller content is clipped). */
  heightRule?: "atLeast" | "exact";
  /** Keep the whole row on one page — never split across a page/column break
   *  (w:cantSplit). */
  cantSplit?: boolean;
  /** Repeat this row as a header at the top of each page the table continues onto
   *  (w:tblHeader). Honored for the leading contiguous header rows. */
  header?: boolean;
}

/** Build the model RowProps a RowOptions describes (undefined when nothing set). */
function rowPropsFrom(o: RowOptions | undefined): RowProps | undefined {
  if (!o) return undefined;
  const props: RowProps = {};
  if (o.height !== undefined && o.height > 0) props.height = { value: o.height, rule: o.heightRule ?? "atLeast" };
  if (o.cantSplit) props.cantSplit = true;
  if (o.header) props.repeatHeader = true;
  return Object.keys(props).length > 0 ? props : undefined;
}

export interface TableOptions {
  /** Column widths as fractions of the content width (normalized to sum 1). */
  colFractions?: number[];
  /** Bold every run in the first row. */
  headerRow?: boolean;
  /** Apply a registered table-style preset (header styling, borders, shading,
   *  striping). Builder-only sugar; explicit per-cell values win. Mutually
   *  exclusive with `styleId` — if both are given, `styleId` wins (with a warning). */
  style?: string;
  /** Reference a REAL table style registered via DocumentBuilder.tableStyle() and
   *  BAKE its effective per-cell formatting onto the cells (so it renders) while
   *  keeping the styleId reference for docx round-trip. Destructive on cell
   *  shading/borders/margin, like applying a style in Word. */
  styleId?: string;
  /** Which conditional bands of the referenced `styleId` are active (w:tblLook).
   *  Default: header row + row banding. */
  condOverrides?: TableCondOverrides;
  /** Column sizing strategy (w:tblLayout): "fixed" (default), "autofitContents",
   *  or "autofitWindow". */
  widthMode?: "fixed" | "autofitContents" | "autofitWindow";
  /** Table-level default borders (OOXML w:tblPr/w:tblBorders), including the two
   *  interior edges (insideH/insideV). Re-emitted at tblPr level on export and
   *  cascaded onto cells that don't set their own (issue #48). */
  borders?: TableBorders;
  /** Table-level default shading fill (w:tblPr/w:shd) — a CSS color applied to every
   *  cell unless the cell overrides it. */
  shading?: string;
  /** Table-level default cell margins (w:tblPr/w:tblCellMar), the base each cell's
   *  own margin overrides per side. */
  cellMargin?: CellMargin;
  /** Table indent from the leading content edge (w:tblPr/w:tblInd), in px. Shifts
   *  the whole table right like a paragraph's left indent. Absent = 0. */
  indent?: number;
  /** Render the table's columns right-to-left (w:tblPr/w:bidiVisual): grid column 0
   *  paints at the right edge. Absent = left-to-right. */
  bidiVisual?: boolean;
  /** Floating-table overlap behavior (w:tblPr/w:tblOverlap). Absent = "overlap"
   *  (Word's default). */
  overlap?: "never" | "overlap";
  /** Table caption / title (w:tblPr/w:tblCaption) — accessibility metadata. */
  caption?: string;
  /** Table description / alt text (w:tblPr/w:tblDescription) — accessibility metadata. */
  description?: string;
}

/** Cell paragraphs are compact (no after-spacing, tighter leading) — matching
 *  how the editor's own table insertion styles cell content. */
const CELL_PARA: Partial<ParaStyle> = { spaceAfterPx: 0, lineHeight: 1.35 };

export class TableBuilder {
  private readonly tableRows: TableRow[] = [];
  private fractions: number[] | undefined;

  constructor(
    private readonly ctx: BuilderContext,
    private readonly opts: TableOptions = {},
  ) {
    this.fractions = opts.colFractions;
  }

  row(cells: CellContent[], rowOpts?: RowOptions): this;
  row(build: (r: RowBuilder) => void, rowOpts?: RowOptions): this;
  row(arg: CellContent[] | ((r: RowBuilder) => void), rowOpts?: RowOptions): this {
    const r = new RowBuilder(this.ctx);
    if (typeof arg === "function") arg(r);
    else for (const c of arg) r.cell(c);
    const row: TableRow = { cells: r.cells };
    const props = rowPropsFrom(rowOpts);
    if (props) row.props = props;
    this.tableRows.push(row);
    return this;
  }

  rows(data: CellContent[][]): this {
    for (const row of data) this.row(row);
    return this;
  }

  colFractions(fractions: number[]): this {
    this.fractions = fractions;
    return this;
  }

  /** Materialize the TableBlock (called by the owning scope's .table()). */
  toBlock(): TableBlock {
    if (this.tableRows.length === 0) {
      this.ctx.warn("table-empty", "A table was built with no rows — a single empty cell was inserted.");
      this.row([""]);
    }
    // `style` (builder preset) and `styleId` (real table style) are mutually
    // exclusive — a real style reference takes precedence; the preset is ignored.
    if (this.opts.style && this.opts.styleId !== undefined) {
      this.ctx.warn(
        "table-style-conflict",
        "Both `style` (preset) and `styleId` (real table style) were passed to .table() — `styleId` wins; the preset was ignored.",
      );
    }
    const preset: TableStylePreset | undefined =
      this.opts.styleId === undefined && this.opts.style ? this.ctx.tableStyle(this.opts.style) : undefined;
    const headerRow = this.opts.headerRow ?? preset?.headerRow ?? false;
    // Apply header styling, borders, shading and striping. Explicit per-cell values
    // (cell.shading/borders set via CellSpec) always win; the preset is the base.
    this.tableRows.forEach((row, ri) => {
      const isHeader = headerRow && ri === 0;
      for (const cell of row.cells) {
        if (isHeader) {
          const hc: Partial<CharStyle> = { bold: true, ...preset?.headerChar };
          for (const block of cell.blocks) {
            if (block.kind === "paragraph") for (const run of block.runs) Object.assign(run.style, hc);
          }
        }
        if (cell.shading === undefined) {
          const sh = isHeader
            ? preset?.headerShading
            : (ri % 2 === 1 ? preset?.stripeShading : undefined) ?? preset?.shading;
          if (sh !== undefined) cell.shading = sh;
        }
        if (cell.borders === undefined && preset?.borders) cell.borders = preset.borders;
      }
    });
    const table: TableBlock = { kind: "table", id: this.ctx.ids.next(), revision: 0, rows: this.tableRows };
    if (this.fractions && this.fractions.length > 0) {
      const sum = this.fractions.reduce((a, b) => a + b, 0);
      if (sum > 0) table.colFractions = this.fractions.map((f) => f / sum);
    }
    if (this.opts.widthMode && this.opts.widthMode !== "fixed") table.widthMode = this.opts.widthMode;
    // A REAL table-style reference: bake the style's effective per-cell formatting
    // onto the cells (the layout engine reads concrete props) and keep the
    // styleId + active bands for docx round-trip. Shares the editor's bake path.
    if (this.opts.styleId !== undefined) this.applyTableStyleRef(table, this.opts.styleId, this.opts.condOverrides);
    this.applyTableDefaults(table);
    // Minor & advanced table props (issue #61).
    if (this.opts.indent !== undefined && this.opts.indent !== 0) table.indentPx = this.opts.indent;
    if (this.opts.bidiVisual) table.bidiVisual = true;
    if (this.opts.overlap !== undefined) table.overlap = this.opts.overlap;
    if (this.opts.caption !== undefined) table.caption = this.opts.caption;
    if (this.opts.description !== undefined) table.description = this.opts.description;
    return table;
  }

  /** Record the table-level defaults (w:tblBorders/w:shd/w:tblCellMar) on the model
   *  so export re-emits them at tblPr level (issue #48), and cascade borders +
   *  shading + cell-margin onto cells that didn't set their own — the layout engine
   *  reads concrete per-cell props, so this makes the table-wide defaults render
   *  headlessly while explicit per-cell values still win. */
  private applyTableDefaults(table: TableBlock): void {
    const { borders, shading, cellMargin } = this.opts;
    if (borders) table.defaultBorders = borders;
    if (shading !== undefined) table.defaultShading = shading;
    if (cellMargin) table.defaultCellMargin = cellMargin;
    if (!borders && shading === undefined && !cellMargin) return;
    for (const row of table.rows) {
      for (const cell of row.cells) {
        if (shading !== undefined && cell.shading === undefined) cell.shading = shading;
        if (cellMargin && cell.margin === undefined) cell.margin = { ...cellMargin };
      }
    }
    if (borders) this.cascadeBordersOntoCells(table, borders);
  }

  /** Resolve the table-level border box onto each cell that set no borders of its
   *  own, mirroring Word's cascade: a cell edge on the table boundary takes the
   *  matching OUTER edge (top/left/bottom/right); an edge between cells takes the
   *  INTERIOR edge (insideH/insideV). Cells are placed on an occupancy grid so
   *  colSpan/rowSpan land on the right boundaries. */
  private cascadeBordersOntoCells(table: TableBlock, b: TableBorders): void {
    const rows = table.rows;
    const width = Math.max(1, ...rows.map((r) => r.cells.reduce((s, c) => s + (c.colSpan ?? 1), 0)));
    const lastRow = rows.length - 1;
    const occupied: boolean[][] = rows.map(() => new Array<boolean>(width).fill(false));
    const ensureRow = (ri: number): void => { while (occupied.length <= ri) occupied.push(new Array<boolean>(width).fill(false)); };
    rows.forEach((row, ri) => {
      let col = 0;
      for (const cell of row.cells) {
        while (col < width && occupied[ri]![col]) col++;
        if (col >= width) break;
        const cs = Math.max(1, cell.colSpan ?? 1);
        const rs = Math.max(1, cell.rowSpan ?? 1);
        const startCol = col;
        const endCol = Math.min(width - 1, col + cs - 1);
        const endRow = ri + rs - 1;
        for (let r = ri; r <= endRow; r++) { ensureRow(r); for (let c = startCol; c <= endCol; c++) occupied[r]![c] = true; }
        if (cell.borders === undefined) {
          const resolved: CellBorders = {};
          const top = ri === 0 ? b.top : b.insideH;
          const bottom = endRow >= lastRow ? b.bottom : b.insideH;
          const left = startCol === 0 ? b.left : b.insideV;
          const right = endCol === width - 1 ? b.right : b.insideV;
          if (top) resolved.top = top;
          if (bottom) resolved.bottom = bottom;
          if (left) resolved.left = left;
          if (right) resolved.right = right;
          if (Object.keys(resolved).length > 0) cell.borders = resolved;
        }
        col = endCol + 1;
      }
    });
  }

  private applyTableStyleRef(table: TableBlock, styleId: string, overrides?: TableCondOverrides): void {
    const look = overrides ?? DEFAULT_TBL_LOOK;
    const styles = this.ctx.doc.tableStyles ?? {};
    const style = styles[styleId];
    table.styleId = styleId;
    table.condOverrides = look;
    if (!style) {
      this.ctx.warn(
        `table-style-ref-missing:${styleId}`,
        `Table style "${styleId}" is not registered (DocumentBuilder.tableStyle) — the reference was set but no formatting was baked.`,
      );
      return;
    }
    // Pass the builder's defaults so the band's char/para also render headlessly
    // (e.g. a header band's bold/white text) — applied only where a cell property
    // was not explicitly set, so an author's explicit CellSpec.style is preserved.
    // The explicit-key maps carry per-cell author provenance so an explicit value
    // that equals the default still survives a conflicting band (#30).
    table.rows = bakeTableStyleRows(table, style, styles, look, {
      char: this.ctx.charDefault,
      para: this.ctx.paraDefault,
      explicitChar: this.ctx.explicitCharKeys,
      explicitPara: this.ctx.explicitParaKeys,
    });
  }
}

export class RowBuilder {
  readonly cells: TableCell[] = [];

  constructor(private readonly ctx: BuilderContext) {}

  cell(content: CellContent, opts?: CellOptions): this;
  cell(build: (s: StoryBuilder) => void, opts?: CellOptions): this;
  cell(content: CellContent | ((s: StoryBuilder) => void), opts?: CellOptions): this {
    const blocks: Block[] = [];
    let spec: CellSpec;
    if (typeof content === "function") {
      content(new StoryBuilder(this.ctx, blocks));
      spec = opts ?? {};
    } else {
      spec = typeof content === "string" ? { ...opts, text: content } : { ...content, ...opts };
      const paraPatch: Partial<ParaStyle> = { ...CELL_PARA };
      if (spec.align !== undefined) paraPatch.align = spec.align;
      const run = this.ctx.run(spec.text ?? "", spec.style ?? {});
      const para = this.ctx.paragraph([run], paraPatch);
      // ctx.run records the author-supplied CharStyle keys (spec.style) on the run as
      // provenance, so table-style baking preserves them even when the value equals the
      // resolved default (#30). Para `align` has no such factory hook — paraPatch also
      // carries builder cell-formatting defaults (CELL_PARA), which must keep the
      // value-equality fallback — so only the author-set `align` is tracked here.
      if (spec.align !== undefined) this.ctx.explicitParaKeys.set(para, new Set(["align"]));
      blocks.push(para);
    }
    // A cell needs at least one paragraph — it is the editor's caret target.
    if (blocks.length === 0) blocks.push(this.ctx.paragraph([], CELL_PARA));
    const cell: TableCell = { id: this.ctx.ids.next(), blocks };
    if (spec.colSpan !== undefined && spec.colSpan > 1) cell.colSpan = spec.colSpan;
    if (spec.rowSpan !== undefined && spec.rowSpan > 1) cell.rowSpan = spec.rowSpan;
    if (spec.shading !== undefined) cell.shading = spec.shading;
    if (spec.borders !== undefined) cell.borders = spec.borders;
    if (spec.margin !== undefined) cell.margin = spec.margin;
    if (spec.preferredWidth !== undefined) cell.preferredWidth = spec.preferredWidth;
    if (spec.vAlign !== undefined && spec.vAlign !== "top") cell.vAlign = spec.vAlign;
    if (spec.textDirection !== undefined && spec.textDirection !== "lrTb") cell.textDirection = spec.textDirection;
    if (spec.noWrap) cell.noWrap = true;
    if (spec.fitText) cell.fitText = true;
    if (spec.hideMark) cell.hideMark = true;
    this.cells.push(cell);
    return this;
  }
}
