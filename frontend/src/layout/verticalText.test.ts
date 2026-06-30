// Vertical cell text (issue #100, OOXML w:textDirection tbRl/btLr). A rotated cell
// lays its text out in a swapped frame: the column auto-sizes NARROW (a stack of
// line heights) while the row grows TALL (the wrapped text length), and the cell
// carries a `rotation` descriptor the paint/geometry layers use. These tests cover
// the sizing and the caret/hit-test mapping through that rotation.
import "./test-canvas-setup";

import { describe, it, expect } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, SectionProps, TableBlock, TableCell } from "@cw/shared";
import { createLayoutEngine } from "./engine";
import { caretRect, hitTest } from "./geometry";
import type { PlacedTable, PlacedTableCell } from "./layoutTree";

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
const table = (rows: TableCell[][], widthMode?: TableBlock["widthMode"]): TableBlock => ({
  kind: "table", id: fresh(), revision: 0, rows: rows.map((cells) => ({ cells })),
  ...(widthMode ? { widthMode } : {}),
});
const doc = (blocks: TableBlock[]): Document => ({ section: SECTION, blocks });

const layoutTable = (t: TableBlock): PlacedTable => {
  for (const pg of createLayoutEngine().layout(doc([t])).pages) {
    for (const pb of pg.blocks) if (pb.blockId === t.id && pb.table) return pb.table;
  }
  throw new Error("table not placed");
};

const HEADER = "Standings";

describe("engine — vertical cell text sizing (w:textDirection)", () => {
  it("auto-sizes a tbRl header column NARROW and its row TALL vs. horizontal", () => {
    const vertical = table([[cell(HEADER, { textDirection: "tbRl" }), cell("alpha")]], "autofitContents");
    const horizontal = table([[cell(HEADER), cell("alpha")]], "autofitContents");

    const vt = layoutTable(vertical);
    const ht = layoutTable(horizontal);

    const vHeaderCol = vt.colWidths[0]!;
    const hHeaderCol = ht.colWidths[0]!;
    // The rotated header column is just a stack of line heights wide — far narrower
    // than the same word laid out horizontally.
    expect(vHeaderCol).toBeLessThan(hHeaderCol);

    const vRow = vt.rows[0]!.height;
    const hRow = ht.rows[0]!.height;
    // ...and the row grows tall: the rotated text length now drives the row height.
    expect(vRow).toBeGreaterThan(hRow);
    // The vertical header is taller than it is wide (the whole point of rotation).
    expect(vRow).toBeGreaterThan(vHeaderCol);
  });

  it("records a +90° (CW) rotation for tbRl and −90° (CCW) for btLr", () => {
    const cw = layoutTable(table([[cell(HEADER, { textDirection: "tbRl" }), cell("x")]]));
    const ccw = layoutTable(table([[cell(HEADER, { textDirection: "btLr" }), cell("x")]]));

    const cwCell = cw.rows[0]!.cells[0]!;
    const ccwCell = ccw.rows[0]!.cells[0]!;
    expect(cwCell.rotation?.angle).toBeCloseTo(Math.PI / 2, 6);
    expect(ccwCell.rotation?.angle).toBeCloseTo(-Math.PI / 2, 6);
    // A plain horizontal cell carries no rotation.
    expect(cw.rows[0]!.cells[1]!.rotation).toBeUndefined();
  });

  it("places the rotation origin on the cell's inner edge", () => {
    const t = layoutTable(table([[cell(HEADER, { textDirection: "tbRl" }), cell("x")]]));
    const c = t.rows[0]!.cells[0]!;
    const rot = c.rotation!;
    // tbRl anchors the TOP-RIGHT inner corner: origin y is at the cell top, origin x
    // is at/inside the cell's right edge.
    expect(rot.originY).toBeGreaterThanOrEqual(c.y);
    expect(rot.originX).toBeLessThanOrEqual(c.x + c.width + 0.01);
    expect(rot.originX).toBeGreaterThan(c.x);
  });
});

describe("geometry — caret/hit-test inside a rotated cell", () => {
  const build = (dir: TableCell["textDirection"]) => {
    const t = table([[cell(HEADER, { textDirection: dir }), cell("alpha")]], "autofitContents");
    const tree = createLayoutEngine().layout(doc([t]));
    let placed: PlacedTableCell | undefined;
    for (const pg of tree.pages) for (const pb of pg.blocks) if (pb.table) placed = pb.table.rows[0]!.cells[0]!;
    if (!placed) throw new Error("cell not placed");
    return { tree, cell: placed, paraId: t.rows[0]!.cells[0]!.blocks[0]!.id };
  };

  it("places the caret for offset 0 inside the rotated cell box", () => {
    for (const dir of ["tbRl", "btLr"] as const) {
      const { tree, cell: c, paraId } = build(dir);
      const r = caretRect(tree, { blockId: paraId, offset: 0 });
      expect(r).not.toBeNull();
      // The caret's center is the insertion point — it must land inside the cell box
      // (the 2px DOM bar is centered on it, so it may straddle the edge by half a
      // line height; orientation stays vertical — a documented limitation).
      const cx = r!.x;
      const cy = r!.y + r!.height / 2;
      expect(cx).toBeGreaterThanOrEqual(c.x - 1);
      expect(cx).toBeLessThanOrEqual(c.x + c.width + 1);
      expect(cy).toBeGreaterThanOrEqual(c.y - 1);
      expect(cy).toBeLessThanOrEqual(c.y + c.height + 1);
    }
  });

  it("hit-tests a click in the rotated cell back to its own paragraph", () => {
    const { tree, cell: c, paraId } = build("tbRl");
    const pos = hitTest(tree, 0, c.x + c.width / 2, c.y + c.height / 2);
    expect(pos).not.toBeNull();
    expect(pos!.blockId).toBe(paraId);
  });

  it("maps the caret DOWN the column as the offset advances (tbRl reads top→bottom)", () => {
    const { tree, paraId } = build("tbRl");
    const first = caretRect(tree, { blockId: paraId, offset: 1 })!;
    const later = caretRect(tree, { blockId: paraId, offset: HEADER.length })!;
    // tbRl maps the text-advance axis to page-y, so a later offset sits lower.
    expect(later.y).toBeGreaterThan(first.y);
  });
});
