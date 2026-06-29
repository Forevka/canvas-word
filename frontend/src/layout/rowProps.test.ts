// Table row properties (w:trPr) at the layout level: a fixed/exact row height
// (w:trHeight), a repeating header row re-drawn atop each continuation page
// (w:tblHeader), and a cant-split row that moves whole across a page break.
import "./test-canvas-setup";

import { describe, it, expect } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, SectionProps, TableBlock, TableCell, TableRow } from "@cw/shared";
import { createLayoutEngine } from "./engine";

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
const fresh = (): string => `r${nextId++}`;
const para = (text: string): Paragraph => ({
  kind: "paragraph", id: fresh(), revision: 0, runs: [{ text, style: { ...CHAR } }], style: { ...PARA },
});
const cell = (text: string): TableCell => ({ id: fresh(), blocks: [para(text)] });
const row = (cells: TableCell[], props?: TableRow["props"]): TableRow => ({ cells, ...(props ? { props } : {}) });
const table = (rows: TableRow[]): TableBlock => ({ kind: "table", id: fresh(), revision: 0, rows });
const doc = (t: TableBlock): Document => ({ section: SECTION, blocks: [t] });

const firstTable = (pb: { table?: unknown }) => pb.table as
  | { rows: { height: number; cells: { originRow: number }[] }[]; firstLineIndex?: number }
  | undefined;

/** Every placed chunk of `id` across all pages, tagged with its page index. */
const placedChunks = (d: Document, id: string): { page: number; rows: { height: number; cells: { originRow: number }[] }[] }[] => {
  const tree = createLayoutEngine().layout(d);
  const out: { page: number; rows: { height: number; cells: { originRow: number }[] }[] }[] = [];
  tree.pages.forEach((pg, pi) => {
    for (const pb of pg.blocks) {
      const t = firstTable(pb);
      if (pb.blockId === id && t) out.push({ page: pi, rows: t.rows });
    }
  });
  return out;
};

describe("engine — table row properties (w:trPr)", () => {
  it("pins a row to an EXACT height even when its content is shorter (w:trHeight exact)", () => {
    const t = table([
      row([cell("tall"), cell("row")], { height: { value: 120, rule: "exact" } }),
      row([cell("normal"), cell("row")]),
    ]);
    const chunks = placedChunks(doc(t), t.id);
    expect(chunks).toHaveLength(1);
    expect(chunks[0]!.rows[0]!.height).toBeCloseTo(120, 0);
    // The second row keeps its content-derived (much smaller) height.
    expect(chunks[0]!.rows[1]!.height).toBeLessThan(120);
  });

  it("grows a row to an at-least floor but never shrinks below content (w:trHeight atLeast)", () => {
    const grown = table([row([cell("x")], { height: { value: 90, rule: "atLeast" } })]);
    const tiny = table([row([cell("x")], { height: { value: 1, rule: "atLeast" } })]);
    const grownH = placedChunks(doc(grown), grown.id)[0]!.rows[0]!.height;
    const tinyH = placedChunks(doc(tiny), tiny.id)[0]!.rows[0]!.height;
    expect(grownH).toBeCloseTo(90, 0); // grown to the floor
    expect(tinyH).toBeGreaterThan(1); // 1px floor is below content → content wins
    expect(tinyH).toBeLessThan(90);
  });

  it("repeats the header row at the top of each continuation page (w:tblHeader)", () => {
    // Header + 20 data rows, each pinned to 60px → spills onto a second page.
    const rows: TableRow[] = [row([cell("Header A"), cell("Header B")], { repeatHeader: true, height: { value: 60, rule: "exact" } })];
    for (let i = 0; i < 20; i++) rows.push(row([cell(`r${i}`), cell("data")], { height: { value: 60, rule: "exact" } }));
    const t = table(rows);
    const chunks = placedChunks(doc(t), t.id);

    // Must paginate.
    const lastPage = Math.max(...chunks.map((c) => c.page));
    expect(lastPage).toBeGreaterThan(0);

    // A continuation page carries a SEPARATE single-row chunk that is the repeated
    // header (its only cell maps back to the model's header row, originRow 0).
    const repeated = chunks.filter(
      (c) => c.page > 0 && c.rows.length === 1 && c.rows[0]!.cells[0]!.originRow === 0,
    );
    expect(repeated.length).toBeGreaterThanOrEqual(1);

    // The data chunk on that same page starts AFTER the header band (originRow > 0).
    const dataOnPage2 = chunks.find((c) => c.page === 1 && c.rows.length > 1);
    expect(dataOnPage2).toBeDefined();
    expect(dataOnPage2!.rows[0]!.cells[0]!.originRow).toBeGreaterThan(0);
  });

  it("moves a cant-split row whole to the next page when it does not fit (w:cantSplit)", () => {
    const t = table([
      row([cell("A")], { height: { value: 760, rule: "exact" } }),
      row([cell("B")], { height: { value: 200, rule: "exact" }, cantSplit: true }),
    ]);
    const chunks = placedChunks(doc(t), t.id);
    // Row B (originRow 1) is placed exactly once, on a later page, at its full height.
    const bPlacements = chunks.flatMap((c) =>
      c.rows.filter((r) => r.cells[0]!.originRow === 1).map((r) => ({ page: c.page, height: r.height })),
    );
    expect(bPlacements).toHaveLength(1);
    expect(bPlacements[0]!.page).toBeGreaterThan(0);
    expect(bPlacements[0]!.height).toBeCloseTo(200, 0);
  });
});
