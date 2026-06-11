// hitTestCell / cellRangeRects: grid-coordinate mapping from page pixels and back.
import "./test-canvas-setup";

import { describe, it, expect } from "vitest";
import type { Block, CharStyle, Document, ParaStyle, Paragraph, SectionProps, TableBlock, TableCell } from "@cw/shared";
import { createLayoutEngine } from "./engine";
import { cellRangeRects, hitTest, hitTestCell } from "./geometry";
import type { PlacedBlock } from "./layoutTree";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA: ParaStyle = { align: "left", lineHeight: 1, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

let id = 0;
const fresh = () => `t${id++}`;
const para = (text: string): Paragraph => ({ kind: "paragraph", id: fresh(), revision: 0, runs: [{ text, style: CHAR }], style: PARA });
const cell = (text: string, extra: Partial<TableCell> = {}): TableCell => ({ id: fresh(), blocks: [para(text)], ...extra });
const table = (rows: TableCell[][]): TableBlock => ({ kind: "table", id: "tbl", revision: 0, rows: rows.map((cells) => ({ cells })) });
const doc = (blocks: Block[]): Document => ({ section: SECTION, blocks });
const layout = (d: Document) => createLayoutEngine().layout(d);

const placedTable = (tree: ReturnType<typeof layout>, tableId: string): { page: number; pb: PlacedBlock } => {
  for (const pg of tree.pages) for (const pb of pg.blocks) if (pb.blockId === tableId && pb.table) return { page: pg.index, pb };
  throw new Error("table not placed");
};

describe("hitTestCell", () => {
  it("maps a point in each cell to its grid coordinates", () => {
    const t = table([[cell("a"), cell("b")], [cell("c"), cell("d")]]);
    const tree = layout(doc([t]));
    const { page, pb } = placedTable(tree, "tbl");
    const pt = pb.table!;
    const center = (r: number, c: number) => {
      const row = pt.rows[r]!;
      let x = pt.x;
      for (let i = 0; i < c; i++) x += pt.colWidths[i]!;
      return { x: x + pt.colWidths[c]! / 2, y: row.y + row.height / 2 };
    };
    for (let r = 0; r < 2; r++)
      for (let c = 0; c < 2; c++) {
        const p = center(r, c);
        expect(hitTestCell(tree, page, p.x, p.y)).toEqual({ tableId: "tbl", row: r, col: c });
      }
  });

  it("returns null off the table", () => {
    const t = table([[cell("a"), cell("b")]]);
    const tree = layout(doc([t]));
    const { page, pb } = placedTable(tree, "tbl");
    expect(hitTestCell(tree, page, pb.table!.x - 20, pb.table!.y - 20)).toBeNull();
  });

  it("maps a rowSpan column past the merge hole", () => {
    // row0: [A(rowSpan2), B] ; row1: [C at grid col 1]
    const t = table([[cell("A", { rowSpan: 2 }), cell("B")], [cell("C")]]);
    const tree = layout(doc([t]));
    const { page, pb } = placedTable(tree, "tbl");
    const pt = pb.table!;
    // A point in row 1, right column → grid (1,1), not (1,0).
    const row1 = pt.rows[1]!;
    const x = pt.x + pt.colWidths[0]! + pt.colWidths[1]! / 2;
    expect(hitTestCell(tree, page, x, row1.y + row1.height / 2)).toEqual({ tableId: "tbl", row: 1, col: 1 });
  });
});

describe("hitTest inside table cells", () => {
  it("pins the caret to the clicked cell even when a neighbour's text is closer", () => {
    // 3x3 grid; the middle-left cell holds one short character while its right
    // neighbour is full of text. Clicking the empty space past the short text
    // used to land in the long neighbour (its text was horizontally closer).
    const t = table([
      [cell("aaaaaaaaaa"), cell("bbbbbbbbbb"), cell("cccccccccc")],
      [cell("x"), cell("yyyyyyyyyyyyyyyyyyyy"), cell("zzzzzzzzzz")],
      [cell("dddddddddd"), cell("eeeeeeeeee"), cell("ffffffffff")],
    ]);
    const tree = layout(doc([t]));
    const { page, pb } = placedTable(tree, "tbl");
    const pt = pb.table!;
    const short = pt.rows[1]!.cells[0]!; // the "x" cell
    const expectedBlockId = short.blocks[0]!.blockId;
    // Click near the right edge of the short cell, past where "x" ends.
    const x = short.x + short.width - 4;
    const y = short.y + short.height / 2;
    const pos = hitTest(tree, page, x, y);
    expect(pos?.blockId).toBe(expectedBlockId);
  });
});

describe("cellRangeRects", () => {
  it("covers the full table for a 0,0..1,1 range", () => {
    const t = table([[cell("a"), cell("b")], [cell("c"), cell("d")]]);
    const tree = layout(doc([t]));
    const { page, pb } = placedTable(tree, "tbl");
    const pt = pb.table!;
    const rects = cellRangeRects(tree, "tbl", 0, 0, 1, 1);
    expect(rects.length).toBe(1);
    expect(rects[0]!.pageIndex).toBe(page);
    expect(rects[0]!.x).toBeCloseTo(pt.x, 1);
    expect(rects[0]!.width).toBeCloseTo(pt.width, 1);
    expect(rects[0]!.height).toBeCloseTo(pt.height, 1);
  });

  it("covers a single column for a 0,0..1,0 range", () => {
    const t = table([[cell("a"), cell("b")], [cell("c"), cell("d")]]);
    const tree = layout(doc([t]));
    const { pb } = placedTable(tree, "tbl");
    const pt = pb.table!;
    const rects = cellRangeRects(tree, "tbl", 0, 0, 1, 0);
    expect(rects[0]!.x).toBeCloseTo(pt.x, 1);
    expect(rects[0]!.width).toBeCloseTo(pt.colWidths[0]!, 1);
  });
});
