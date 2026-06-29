// hitTestRowBoundary: page-pixel → row-boundary mapping that drives the
// drag-to-resize row-height interaction (#76), mirroring the column-boundary path.
import "./test-canvas-setup";

import { describe, it, expect } from "vitest";
import type { Block, CharStyle, Document, ParaStyle, Paragraph, SectionProps, TableBlock, TableCell, TableRow } from "@cw/shared";
import { createLayoutEngine } from "./engine";
import { hitTestColumnBoundary, hitTestRowBoundary } from "./geometry";
import type { PlacedBlock } from "./layoutTree";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA: ParaStyle = { align: "left", lineHeight: 1, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

let id = 0;
const fresh = () => `t${id++}`;
const para = (text: string): Paragraph => ({ kind: "paragraph", id: fresh(), revision: 0, runs: [{ text, style: CHAR }], style: PARA });
const cell = (text: string, extra: Partial<TableCell> = {}): TableCell => ({ id: fresh(), blocks: [para(text)], ...extra });
const row = (cells: TableCell[], props?: TableRow["props"]): TableRow => ({ cells, ...(props ? { props } : {}) });
const table = (rows: TableRow[]): TableBlock => ({ kind: "table", id: "tbl", revision: 0, rows });
const doc = (blocks: Block[]): Document => ({ section: SECTION, blocks });
const layout = (d: Document) => createLayoutEngine().layout(d);

const placedTable = (tree: ReturnType<typeof layout>, tableId: string): { page: number; pb: PlacedBlock } => {
  for (const pg of tree.pages) for (const pb of pg.blocks) if (pb.blockId === tableId && pb.table) return { page: pg.index, pb };
  throw new Error("table not placed");
};

describe("hitTestRowBoundary", () => {
  it("maps a point on each row's bottom edge to that row's model index", () => {
    const t = table([row([cell("a"), cell("b")]), row([cell("c"), cell("d")]), row([cell("e"), cell("f")])]);
    const tree = layout(doc([t]));
    const { page, pb } = placedTable(tree, "tbl");
    const pt = pb.table!;
    for (let r = 0; r < pt.rows.length; r++) {
      const edge = pt.rows[r]!.y + pt.rows[r]!.height;
      const x = pt.x + pt.width / 2;
      const hit = hitTestRowBoundary(tree, page, x, edge);
      expect(hit).not.toBeNull();
      expect(hit!.tableId).toBe("tbl");
      expect(hit!.rowIndex).toBe(r);
      expect(hit!.y).toBeCloseTo(edge, 0);
    }
  });

  it("returns null in a row's interior and off the table", () => {
    const t = table([row([cell("a")]), row([cell("b")])]);
    const tree = layout(doc([t]));
    const { page, pb } = placedTable(tree, "tbl");
    const pt = pb.table!;
    const r0 = pt.rows[0]!;
    expect(hitTestRowBoundary(tree, page, pt.x + pt.width / 2, r0.y + r0.height / 2)).toBeNull();
    expect(hitTestRowBoundary(tree, page, pt.x - 20, r0.y + r0.height)).toBeNull();
  });

  it("prefers the column grip at a cell corner (column tested first)", () => {
    // At the corner where an interior column boundary meets a row's bottom edge,
    // both hit-tests fire; the caller checks columns first so the corner resizes
    // the column, never the row — keeping the two interactions from colliding.
    const t = table([row([cell("a"), cell("b")]), row([cell("c"), cell("d")])]);
    const tree = layout(doc([t]));
    const { page, pb } = placedTable(tree, "tbl");
    const pt = pb.table!;
    const corner = { x: pt.x + pt.colWidths[0]!, y: pt.rows[0]!.y + pt.rows[0]!.height };
    expect(hitTestColumnBoundary(tree, page, corner.x, corner.y)).not.toBeNull();
    expect(hitTestRowBoundary(tree, page, corner.x, corner.y)).not.toBeNull();
  });
});
