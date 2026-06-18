// Table styles must survive a full .docx round-trip (the style with its
// conditional bands, the table's w:tblStyle reference and w:tblLook flags) and
// the editor commands must bake the effective per-cell formatting.
import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, SectionProps, TableBlock, TableCell, TableStyle } from "@cw/shared";
import { applyOp } from "@cw/shared";
import { applyTableStyle, deleteTableStyle, upsertTableStyle } from "../../editor/commands";
import type { Command, EditorState } from "../../editor/state";
import { runImport } from "../../import/docx/pipeline";
import { writeDocx } from "./writeDocx";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 14, bold: false, italic: false, underline: false, strikethrough: false, color: "#202124" };
const PARA = { align: "left" as const, lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

const gridStyle: TableStyle = {
  id: "GridAccent",
  name: "Grid Accent",
  rowBandSize: 1,
  conds: {
    wholeTable: { borders: { top: { color: "#000000", widthPx: 1 }, bottom: { color: "#000000", widthPx: 1 }, left: { color: "#000000", widthPx: 1 }, right: { color: "#000000", widthPx: 1 } } },
    firstRow: { shading: "#4472c4", char: { bold: true, color: "#ffffff" } },
    band1Horz: { shading: "#d9e2f3" },
  },
};

let n = 0;
const cellPara = (text: string): Paragraph => ({ kind: "paragraph", id: `tp${n++}`, revision: 0, runs: [{ text, style: { ...CHAR } }], style: { ...PARA } });
const cell = (text: string): TableCell => ({ id: `tc${n++}`, blocks: [cellPara(text)] });
const mkTable = (styleId?: string): TableBlock => ({
  kind: "table", id: "T1", revision: 0,
  colFractions: [0.5, 0.5],
  rows: [
    { cells: [cell("H1"), cell("H2")] },
    { cells: [cell("a"), cell("b")] },
    { cells: [cell("c"), cell("d")] },
  ],
  ...(styleId ? { styleId, condOverrides: { firstRow: true, bandRows: true } } : {}),
});

const run = (state: EditorState, cmd: Command): Document => {
  const trn = cmd(state);
  if (!trn) throw new Error("no transaction");
  let doc = state.doc;
  for (const op of trn.ops) doc = applyOp(doc, op).doc;
  return doc;
};
const firstCellOf = (t: TableBlock, r: number, c: number): TableCell => t.rows[r]!.cells[c]!;

describe("table style commands", () => {
  it("applyTableStyle bakes header + banding onto the cells and sets the reference", () => {
    const table = mkTable();
    const doc: Document = { section: SECTION, tableStyles: { GridAccent: gridStyle }, blocks: [table] };
    const focusId = (table.rows[0]!.cells[0]!.blocks[0] as Paragraph).id;
    const state: EditorState = { doc, selection: { anchor: { blockId: focusId, offset: 0 }, focus: { blockId: focusId, offset: 0 } } };
    const out = run(state, applyTableStyle("GridAccent"));
    const t = out.blocks[0] as TableBlock;
    expect(t.styleId).toBe("GridAccent");
    expect(firstCellOf(t, 0, 0).shading).toBe("#4472c4"); // header
    expect(firstCellOf(t, 1, 0).shading).toBe("#d9e2f3"); // first body row → band1
    expect(firstCellOf(t, 0, 0).borders?.top).toBeTruthy(); // wholeTable border
  });

  it("upsertTableStyle re-bakes referencing tables", () => {
    const table = mkTable("GridAccent");
    const doc: Document = { section: SECTION, tableStyles: { GridAccent: gridStyle }, blocks: [table] };
    const recolored: TableStyle = { ...gridStyle, conds: { ...gridStyle.conds, firstRow: { shading: "#ff0000" } } };
    const out = run({ doc, selection: null }, upsertTableStyle(recolored));
    expect(firstCellOf(out.blocks[0] as TableBlock, 0, 0).shading).toBe("#ff0000");
  });

  it("deleteTableStyle drops the reference but keeps the baked cells", () => {
    const table = mkTable("GridAccent");
    table.rows[0]!.cells[0]!.shading = "#4472c4";
    const doc: Document = { section: SECTION, tableStyles: { GridAccent: gridStyle }, blocks: [table] };
    const out = run({ doc, selection: null }, deleteTableStyle("GridAccent"));
    expect(out.tableStyles?.GridAccent).toBeUndefined();
    const t = out.blocks[0] as TableBlock;
    expect(t.styleId).toBeUndefined();
    expect(firstCellOf(t, 0, 0).shading).toBe("#4472c4"); // baked formatting kept
  });
});

describe("table style .docx round-trip", () => {
  it("re-imports the table style (conds), the w:tblStyle ref and w:tblLook flags", () => {
    const doc: Document = { section: SECTION, tableStyles: { GridAccent: gridStyle }, blocks: [mkTable("GridAccent")] };
    const bytes = writeDocx(doc).bytes;
    const back = runImport(bytes).doc;

    const st = back.tableStyles?.GridAccent;
    expect(st).toBeTruthy();
    expect(st!.conds.firstRow?.shading).toBe("#4472c4");
    expect(st!.conds.band1Horz?.shading).toBe("#d9e2f3");
    expect(st!.conds.wholeTable?.borders?.top).toBeTruthy();

    const t = back.blocks.find((b): b is TableBlock => b.kind === "table")!;
    expect(t.styleId).toBe("GridAccent");
    expect(t.condOverrides?.firstRow).toBe(true);
    expect(t.condOverrides?.bandRows).toBe(true);
  });
});
