// Table composition. Two entry shapes share this implementation:
//   - data-driven: .table([["Name", "DOB"], ["Ada", "1815"]], { headerRow: true })
//   - structural:  .table(t => t.row(r => r.cell("x", { colSpan: 2 })))
// Cells hold full block stories; a cell callback gets a StoryBuilder, so
// paragraphs/images/lists inside cells reuse the normal scope surface.

import type { Block, CellBorders, CellMargin, CharStyle, ParaStyle, TableBlock, TableCell, TableRow } from "@cw/shared";
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
}

export type CellContent = string | CellSpec;

export type CellOptions = Omit<CellSpec, "text">;

export interface TableOptions {
  /** Column widths as fractions of the content width (normalized to sum 1). */
  colFractions?: number[];
  /** Bold every run in the first row. */
  headerRow?: boolean;
  /** Apply a registered table-style preset (header styling, borders, shading,
   *  striping). Builder-only sugar; explicit per-cell values win. */
  style?: string;
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
    const preset: TableStylePreset | undefined = this.opts.style ? this.ctx.tableStyle(this.opts.style) : undefined;
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
    return table;
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
      blocks.push(this.ctx.paragraph([this.ctx.run(spec.text ?? "", spec.style ?? {})], paraPatch));
    }
    // A cell needs at least one paragraph — it is the editor's caret target.
    if (blocks.length === 0) blocks.push(this.ctx.paragraph([], CELL_PARA));
    const cell: TableCell = { id: this.ctx.ids.next(), blocks };
    if (spec.colSpan !== undefined && spec.colSpan > 1) cell.colSpan = spec.colSpan;
    if (spec.rowSpan !== undefined && spec.rowSpan > 1) cell.rowSpan = spec.rowSpan;
    if (spec.shading !== undefined) cell.shading = spec.shading;
    if (spec.borders !== undefined) cell.borders = spec.borders;
    if (spec.margin !== undefined) cell.margin = spec.margin;
    this.cells.push(cell);
    return this;
  }
}
