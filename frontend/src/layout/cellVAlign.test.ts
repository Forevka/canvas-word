// Cell vertical alignment (w:vAlign). A short paragraph in a cell that is taller
// than its content (forced tall by a sibling row or a rowSpan) must sit at the
// top (default), centered, or hugging the bottom per the cell's vAlign.
import "./test-canvas-setup";

import { describe, it, expect } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, SectionProps, TableBlock, TableCell } from "@cw/shared";
import { createLayoutEngine } from "./engine";
import type { PlacedBlock } from "./layoutTree";

const CHAR: CharStyle = {
  fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false,
  underline: false, strikethrough: false, color: "#000000",
};
const PARA: ParaStyle = {
  align: "left", lineHeight: 1, spaceBeforePx: 0, spaceAfterPx: 0,
  indentFirstLinePx: 0, indentLeftPx: 0,
};
const SECTION: SectionProps = {
  pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
};

let nextId = 0;
const fresh = (): string => `v${nextId++}`;
const para = (text: string): Paragraph => ({
  kind: "paragraph", id: fresh(), revision: 0, runs: [{ text, style: { ...CHAR } }], style: { ...PARA },
});
const cell = (text: string, extra: Partial<TableCell> = {}): TableCell => ({ id: fresh(), blocks: [para(text)], ...extra });
const table = (rows: TableCell[][], colFractions?: number[]): TableBlock => ({
  kind: "table", id: fresh(), revision: 0, rows: rows.map((cells) => ({ cells })),
  ...(colFractions ? { colFractions } : {}),
});
const doc = (blocks: TableBlock[]): Document => ({ section: SECTION, blocks });

const layout = (d: Document) => createLayoutEngine().layout(d);
const placedTable = (d: Document, id: string) => {
  for (const pg of layout(d).pages) for (const pb of pg.blocks) if (pb.blockId === id && pb.table) return pb.table;
  throw new Error("table not placed");
};
// PlacedTableCell exposes the cell box plus its placed content blocks.
type PlacedCell = { x: number; y: number; height: number; blocks: PlacedBlock[] };

// A multi-line filler that wraps across several lines so its column is much taller
// than a single short line — that height is what the short neighbour aligns within.
const TALL = "wwww wwww wwww wwww wwww wwww wwww wwww wwww wwww wwww wwww wwww wwww";

describe("engine — cell vertical alignment (w:vAlign)", () => {
  it("offsets a short cell's content for center/bottom relative to top within a tall row", () => {
    const mk = (v?: TableCell["vAlign"]): TableBlock =>
      table([[cell("hi", v ? { vAlign: v } : {}), cell(TALL)]], [0.5, 0.5]);

    const cellOf = (t: TableBlock): PlacedCell =>
      placedTable(doc([t]), t.id).rows[0]!.cells[0]! as unknown as PlacedCell;

    const top = cellOf(mk("top"));
    const center = cellOf(mk("center"));
    const bottom = cellOf(mk("bottom"));

    const yT = top.blocks[0]!.y;
    const yC = center.blocks[0]!.y;
    const yB = bottom.blocks[0]!.y;

    // top hugs the cell top; center and bottom progressively push the content down.
    expect(yC).toBeGreaterThan(yT);
    expect(yB).toBeGreaterThan(yC);
    // center sits roughly halfway between top and bottom.
    expect(yC).toBeCloseTo((yT + yB) / 2, 0);
  });

  it("centers/bottom-aligns a short paragraph inside a tall rowSpan cell", () => {
    // The short label spans two rows that are forced tall by their siblings, so the
    // rowSpan cell is much taller than its one line of content.
    const mk = (v?: TableCell["vAlign"]): TableBlock =>
      table(
        [
          [cell("label", v ? { rowSpan: 2, vAlign: v } : { rowSpan: 2 }), cell(TALL)],
          [cell(TALL)],
        ],
        [0.4, 0.6],
      );
    const spanOf = (t: TableBlock): PlacedCell =>
      placedTable(doc([t]), t.id).rows[0]!.cells[0]! as unknown as PlacedCell;

    const spanTop = spanOf(mk());
    const spanCenter = spanOf(mk("center"));
    const spanBottom = spanOf(mk("bottom"));

    // The rowSpan cell box (height = both rows) is much taller than its one line.
    expect(spanCenter.height).toBeGreaterThan(spanCenter.blocks[0]!.lines.length * 16 + 16);
    expect(spanBottom.height).toBeCloseTo(spanTop.height, 1);

    const yT = spanTop.blocks[0]!.y;
    const yC = spanCenter.blocks[0]!.y;
    const yB = spanBottom.blocks[0]!.y;

    expect(yC).toBeGreaterThan(yT);
    expect(yB).toBeGreaterThan(yC);
    // bottom-aligned content sits well below the cell's vertical midpoint.
    expect(yB).toBeGreaterThan(spanBottom.y + spanBottom.height / 2);
  });
});
