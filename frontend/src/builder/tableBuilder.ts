// Table composition. Two entry shapes share this implementation:
//   - data-driven: .table([["Name", "DOB"], ["Ada", "1815"]], { headerRow: true })
//   - structural:  .table(t => t.row(r => r.cell("x", { colSpan: 2 })))
// Cells hold full block stories; a cell callback gets a StoryBuilder, so
// paragraphs/images/lists inside cells reuse the normal scope surface.

import type { Block, CellBorders, CellMargin, CharStyle, ParaStyle, TableBlock, TableCell, TableCondOverrides, TableRow } from "@cw/shared";
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
}

export type CellContent = string | CellSpec;

export type CellOptions = Omit<CellSpec, "text">;

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

  row(cells: CellContent[]): this;
  row(build: (r: RowBuilder) => void): this;
  row(arg: CellContent[] | ((r: RowBuilder) => void)): this {
    const r = new RowBuilder(this.ctx);
    if (typeof arg === "function") arg(r);
    else for (const c of arg) r.cell(c);
    this.tableRows.push({ cells: r.cells });
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
    return table;
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
    this.cells.push(cell);
    return this;
  }
}
