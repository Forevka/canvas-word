// Table indent (w:tblInd) and RTL visual column order (w:bidiVisual), issue #61.
// tblInd shifts the whole table right from the content edge; bidiVisual mirrors the
// columns so grid column 0 paints at the RIGHT of the table (its content is
// unchanged — only x placement flips).
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
const fresh = (): string => `m${nextId++}`;
const para = (text: string): Paragraph => ({
  kind: "paragraph", id: fresh(), revision: 0, runs: [{ text, style: { ...CHAR } }], style: { ...PARA },
});
const cell = (text: string): TableCell => ({ id: fresh(), blocks: [para(text)] });
const table = (extra: Partial<TableBlock> = {}): TableBlock => ({
  kind: "table", id: fresh(), revision: 0,
  colFractions: [0.5, 0.5],
  rows: [{ cells: [cell("L"), cell("R")] }],
  ...extra,
});
const doc = (t: TableBlock): Document => ({ section: SECTION, blocks: [t] });

const layout = (d: Document) => createLayoutEngine().layout(d);
type PlacedCell = { x: number; y: number; height: number; blocks: PlacedBlock[] };
const placedTable = (t: TableBlock) => {
  for (const pg of layout(doc(t)).pages) for (const pb of pg.blocks) if (pb.blockId === t.id && pb.table) return pb.table;
  throw new Error("table not placed");
};

describe("engine — table indent (w:tblInd)", () => {
  it("shifts the whole table right by indentPx, leaving column structure intact", () => {
    const plain = placedTable(table());
    const indented = placedTable(table({ indentPx: 40 }));
    const x0Plain = plain.rows[0]!.cells[0]! as unknown as PlacedCell;
    const x0Indent = indented.rows[0]!.cells[0]! as unknown as PlacedCell;
    // The first cell moves right by exactly the indent.
    expect(x0Indent.x - x0Plain.x).toBeCloseTo(40, 1);
    // The two columns keep their relative offset (only the table origin shifted).
    const c1Plain = plain.rows[0]!.cells[1]! as unknown as PlacedCell;
    const c1Indent = indented.rows[0]!.cells[1]! as unknown as PlacedCell;
    expect(c1Indent.x - c1Plain.x).toBeCloseTo(40, 1);
  });
});

describe("engine — RTL visual column order (w:bidiVisual)", () => {
  it("mirrors the columns so grid column 0 sits to the RIGHT of column 1", () => {
    const ltr = placedTable(table());
    const rtl = placedTable(table({ bidiVisual: true }));
    const l0 = ltr.rows[0]!.cells[0]! as unknown as PlacedCell;
    const l1 = ltr.rows[0]!.cells[1]! as unknown as PlacedCell;
    const r0 = rtl.rows[0]!.cells[0]! as unknown as PlacedCell;
    const r1 = rtl.rows[0]!.cells[1]! as unknown as PlacedCell;
    // LTR: cell 0 left of cell 1. RTL: cell 0 right of cell 1.
    expect(l0.x).toBeLessThan(l1.x);
    expect(r0.x).toBeGreaterThan(r1.x);
    // The mirror is symmetric: cell 0's RTL x equals cell 1's LTR x (equal columns).
    expect(r0.x).toBeCloseTo(l1.x, 1);
    expect(r1.x).toBeCloseTo(l0.x, 1);
  });
});
