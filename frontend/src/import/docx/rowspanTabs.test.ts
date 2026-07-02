// Vertical cell merge → rowSpan, and tab stops (w:tabs → ParaStyle.tabStops).

import { describe, expect, it } from "vitest";
import type { Paragraph, TableBlock } from "@cw/shared";
import { runImport } from "./pipeline";
import { simpleDocx } from "./fixture";

const importBody = (body: string) => runImport(simpleDocx(body));
const para = (b: unknown): Paragraph => {
  expect((b as Paragraph).kind).toBe("paragraph");
  return b as Paragraph;
};
const table = (r: ReturnType<typeof runImport>): TableBlock => r.doc.blocks.find((b) => b.kind === "table") as TableBlock;
const cellText = (t: TableBlock, r: number, c: number) =>
  (t.rows[r]!.cells[c]!.blocks[0] as Paragraph).runs.map((x) => x.text).join("");

const tc = (text: string, tcPr = "") => `<w:tc>${tcPr}<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const vRestart = `<w:tcPr><w:vMerge w:val="restart"/></w:tcPr>`;
const vCont = `<w:tcPr><w:vMerge/></w:tcPr>`;

describe("row spans (w:vMerge)", () => {
  it("spans three rows and drops the two continuation cells", () => {
    const r = importBody(
      `<w:tbl>` +
        `<w:tr>${tc("merged", vRestart)}${tc("b1")}</w:tr>` +
        `<w:tr>${tc("", vCont)}${tc("b2")}</w:tr>` +
        `<w:tr>${tc("", vCont)}${tc("b3")}</w:tr>` +
        `</w:tbl>`,
    );
    const t = table(r);
    expect(t.rows[0]!.cells[0]!.rowSpan).toBe(3);
    expect(t.rows[1]!.cells).toHaveLength(1);
    expect(cellText(t, 1, 0)).toBe("b2");
    expect(cellText(t, 2, 0)).toBe("b3");
    // Column count stays 2 even though rows 1-2 carry a single cell each.
    expect(t.colFractions).toHaveLength(2);
  });

  it("keeps the spanning cell's column position via colFractions width", () => {
    // Merge in the SECOND column: continuation row's lone cell must still be
    // understood as column 0, with column 1 held by the span.
    const r = importBody(
      `<w:tbl>` +
        `<w:tr>${tc("a1")}${tc("tall", vRestart)}</w:tr>` +
        `<w:tr>${tc("a2")}${tc("", vCont)}</w:tr>` +
        `</w:tbl>`,
    );
    const t = table(r);
    expect(t.rows[0]!.cells[1]!.rowSpan).toBe(2);
    expect(t.rows[1]!.cells).toHaveLength(1);
    expect(cellText(t, 1, 0)).toBe("a2");
    expect(t.colFractions).toHaveLength(2);
  });

  it("a borderless rowspan cell's bottom is outer only on its last spanned row", () => {
    const borders = `<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>`;
    const r = importBody(
      `<w:tbl>${borders}` +
        `<w:tr>${tc("m", vRestart)}${tc("b1")}</w:tr>` +
        `<w:tr>${tc("", vCont)}${tc("b2")}</w:tr>` +
        `</w:tbl>`,
    );
    const t = table(r);
    // The spanning cell reaches the table bottom → its bottom edge is the outer one.
    expect(t.rows[0]!.cells[0]!.rowSpan).toBe(2);
    expect(t.rows[0]!.cells[0]!.borders!.bottom).toBeDefined();
    expect(t.rows[0]!.cells[0]!.borders!.top).toBeDefined();
  });
});

describe("tab stops (w:tabs)", () => {
  const tabsP = (tabs: string, runs: string) =>
    `<w:p><w:pPr><w:tabs>${tabs}</w:tabs></w:pPr>${runs}</w:p>`;

  it("maps positions (twips→px), alignment, and leaders", () => {
    const r = importBody(
      tabsP(
        `<w:tab w:val="left" w:pos="720"/>` +
          `<w:tab w:val="center" w:pos="2160"/>` +
          `<w:tab w:val="right" w:pos="4320" w:leader="dot"/>` +
          `<w:tab w:val="decimal" w:pos="6480"/>`,
        `<w:r><w:t>x</w:t></w:r>`,
      ),
    );
    expect(para(r.doc.blocks[0]).style.tabStops).toEqual([
      { posPx: 48 }, // left default — align omitted
      { posPx: 144, align: "center" },
      { posPx: 288, align: "right", leader: "dot" },
      { posPx: 432, align: "decimal" },
    ]);
  });

  it("sorts stops by position, preserves clear (issue #154), skips bar", () => {
    const r = importBody(
      tabsP(
        `<w:tab w:val="left" w:pos="2880"/><w:tab w:val="clear" w:pos="1440"/>` +
          `<w:tab w:val="bar" w:pos="100"/><w:tab w:val="left" w:pos="720"/>`,
        `<w:r><w:t>x</w:t></w:r>`,
      ),
    );
    const stops = para(r.doc.blocks[0]).style.tabStops!;
    // bar@100 dropped; clear@1440 (96px) kept as a removal marker between the two live stops.
    expect(stops.map((s) => s.posPx)).toEqual([48, 96, 192]);
    expect(stops.find((s) => s.posPx === 96)!.cleared).toBe(true);
    expect(stops.filter((s) => !s.cleared).map((s) => s.posPx)).toEqual([48, 192]);
  });

  it("preserves the tab character in run text", () => {
    const r = importBody(`<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>`);
    expect(para(r.doc.blocks[0]).runs.map((x) => x.text).join("")).toBe("a\tb");
  });
});
