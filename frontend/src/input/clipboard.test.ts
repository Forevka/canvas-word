// Rectangular table-cell copy: HTML <table> + tab/newline plain text.
import { describe, it, expect } from "vitest";
import type { CharStyle, ParaStyle, Paragraph, TableBlock, TableCell } from "@cw/shared";
import { tableRectToClipboard } from "./clipboard";

const CHAR: CharStyle = {
  fontFamily: "Georgia, serif",
  fontSizePx: 16,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#000000",
};
const PARA: ParaStyle = { align: "left", lineHeight: 1, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };

let n = 0;
const fresh = (): string => `t${n++}`;
const para = (text: string): Paragraph => ({ kind: "paragraph", id: fresh(), revision: 0, runs: [{ text, style: CHAR }], style: PARA });
const cell = (text: string, extra: Partial<TableCell> = {}): TableCell => ({ id: fresh(), blocks: [para(text)], ...extra });
const table = (rows: TableCell[][]): TableBlock => ({ kind: "table", id: "tbl", revision: 0, rows: rows.map((cells) => ({ cells })) });

describe("tableRectToClipboard", () => {
  it("serializes a 2x2 rectangle to tab/newline text and an HTML table", () => {
    const t = table([[cell("A"), cell("B")], [cell("C"), cell("D")]]);
    const { html, text } = tableRectToClipboard(t, { r0: 0, c0: 0, r1: 1, c1: 1 });
    expect(text).toBe("A\tB\nC\tD");
    expect(html.startsWith("<table>")).toBe(true);
    expect((html.match(/<tr>/g) ?? []).length).toBe(2);
    expect((html.match(/<td/g) ?? []).length).toBe(4);
    expect(html).toContain("A");
    expect(html).toContain("D");
  });

  it("normalizes a partial rectangle to whole merged cells and emits colspan", () => {
    // A full-width title (colSpan 3) above a normal 3-column row. Selecting only a
    // sub-column of the title must pull in the whole merged cell.
    const t = table([[cell("Title", { colSpan: 3 })], [cell("a"), cell("b"), cell("c")]]);
    const { html, text } = tableRectToClipboard(t, { r0: 0, c0: 1, r1: 1, c1: 1 });
    expect(text.split("\n")[0]).toContain("Title");
    expect(text.split("\n")[1]).toBe("a\tb\tc");
    expect(html).toContain('colspan="3"');
    expect((html.match(/Title/g) ?? []).length).toBe(1); // merged cell emitted once
  });

  it("blanks continuation columns in plain text to keep tab alignment", () => {
    const t = table([[cell("Title", { colSpan: 2 }), cell("X")], [cell("a"), cell("b"), cell("c")]]);
    const { text } = tableRectToClipboard(t, { r0: 0, c0: 0, r1: 1, c1: 2 });
    expect(text.split("\n")[0]).toBe("Title\t\tX"); // colspan-2 cell, blank continuation, then X
    expect(text.split("\n")[1]).toBe("a\tb\tc");
  });
});
