import { describe, expect, it } from "vitest";
import { DEFAULT_CHAR_STYLE, DEFAULT_PARA_STYLE } from "./defaults";
import type { Paragraph, Run, TableBlock } from "./document";
import { bakeTableStyleRows, cellCondFlags, effectiveCellProps, resolveTableStyle, type TableStyle } from "./tableStyles";

const style: TableStyle = {
  id: "Grid",
  name: "Grid",
  rowBandSize: 1,
  conds: {
    wholeTable: { shading: "#ffffff", borders: { top: { color: "#000", widthPx: 1 } } },
    firstRow: { shading: "#4472c4", char: { bold: true } },
    band1Horz: { shading: "#d9e2f3" },
  },
};
const styles = { Grid: style };

describe("resolveTableStyle", () => {
  it("collapses the basedOn chain, child overriding per condition", () => {
    const base: TableStyle = { id: "Base", name: "Base", conds: { wholeTable: { shading: "#eeeeee", margin: { top: 2, right: 2, bottom: 2, left: 2 } } } };
    const child: TableStyle = { id: "Child", name: "Child", basedOn: "Base", conds: { wholeTable: { shading: "#ffffff" } } };
    const r = resolveTableStyle({ Base: base, Child: child }, "Child");
    expect(r.wholeTable!.shading).toBe("#ffffff"); // child overrides
    expect(r.wholeTable!.margin).toEqual({ top: 2, right: 2, bottom: 2, left: 2 }); // inherited
  });
  it("is cycle-guarded", () => {
    const a: TableStyle = { id: "A", name: "A", basedOn: "B", conds: {} };
    const b: TableStyle = { id: "B", name: "B", basedOn: "A", conds: {} };
    expect(() => resolveTableStyle({ A: a, B: b }, "A")).not.toThrow();
  });
});

describe("effectiveCellProps precedence", () => {
  const resolved = resolveTableStyle(styles, "Grid");
  it("firstRow overrides wholeTable and banding for a header cell", () => {
    const flags = cellCondFlags(0, 0, 4, 3, { firstRow: true, bandRows: true });
    const props = effectiveCellProps(resolved, flags);
    expect(props.shading).toBe("#4472c4");
    expect(props.char?.bold).toBe(true);
  });
  it("banded body row gets band1 shading; wholeTable border still applies", () => {
    // Row 1 is the first body row (header at row 0) → band1.
    const flags = cellCondFlags(1, 0, 4, 3, { firstRow: true, bandRows: true });
    expect(flags.rowBand).toBe(1);
    const props = effectiveCellProps(resolved, flags);
    expect(props.shading).toBe("#d9e2f3");
    expect(props.borders?.top).toBeTruthy(); // wholeTable border merged through
  });
  it("direct cell formatting wins last", () => {
    const flags = cellCondFlags(0, 0, 4, 3, { firstRow: true });
    const props = effectiveCellProps(resolved, flags, { shading: "#ff0000" });
    expect(props.shading).toBe("#ff0000");
  });
});

describe("bakeTableStyleRows content provenance (#30)", () => {
  // A header band that recolors first-row text white — conflicts with a cell whose
  // run was explicitly set back to the default color.
  const band: TableStyle = { id: "Hdr", name: "Hdr", conds: { firstRow: { char: { color: "#ffffff" } } } };
  const styles = { Hdr: band };

  // One header cell holding a single run whose color EQUALS the resolved default.
  function headerTable(): { table: TableBlock; run: Run } {
    const run: Run = { text: "H", style: { ...DEFAULT_CHAR_STYLE } }; // color === default (#202124)
    const para: Paragraph = { kind: "paragraph", id: "p", revision: 0, runs: [run], style: { ...DEFAULT_PARA_STYLE } };
    const table: TableBlock = { kind: "table", id: "t", revision: 0, rows: [{ cells: [{ id: "c", blocks: [para] }] }] };
    return { table, run };
  }

  const bakedRun = (rows: ReturnType<typeof bakeTableStyleRows>): Run =>
    (rows[0]!.cells[0]!.blocks[0] as Paragraph).runs[0]!;

  it("preserves an explicitly-set value that equals the default against a conflicting band", () => {
    const { table, run } = headerTable();
    const explicitChar = new WeakMap<object, ReadonlySet<string>>([[run, new Set(["color"])]]);
    const rows = bakeTableStyleRows(table, band, styles, { firstRow: true }, {
      char: DEFAULT_CHAR_STYLE,
      para: DEFAULT_PARA_STYLE,
      explicitChar,
    });
    // The author pinned the default color explicitly, so the band must NOT win.
    expect(bakedRun(rows).style.color).toBe(DEFAULT_CHAR_STYLE.color);
  });

  it("lets genuinely-default (unset) content pick up the band", () => {
    const { table } = headerTable(); // same default color, but no provenance recorded
    const rows = bakeTableStyleRows(table, band, styles, { firstRow: true }, {
      char: DEFAULT_CHAR_STYLE,
      para: DEFAULT_PARA_STYLE,
      explicitChar: new WeakMap(),
    });
    expect(bakedRun(rows).style.color).toBe("#ffffff");
  });

  // The paragraph branch mirrors the run branch: explicitPara provenance must
  // protect an author-set para field (align) that equals the default.
  const alignBand: TableStyle = { id: "Align", name: "Align", conds: { firstRow: { para: { align: "center" } } } };
  const bakedAlign = (rows: ReturnType<typeof bakeTableStyleRows>): string | undefined =>
    (rows[0]!.cells[0]!.blocks[0] as Paragraph).style.align;

  it("preserves an explicitly-set paragraph align that equals the default against a conflicting band", () => {
    const { table } = headerTable();
    const para = table.rows[0]!.cells[0]!.blocks[0]!; // align === default ("left")
    const explicitPara = new WeakMap<object, ReadonlySet<string>>([[para, new Set(["align"])]]);
    const rows = bakeTableStyleRows(table, alignBand, { Align: alignBand }, { firstRow: true }, {
      char: DEFAULT_CHAR_STYLE,
      para: DEFAULT_PARA_STYLE,
      explicitPara,
    });
    expect(bakedAlign(rows)).toBe(DEFAULT_PARA_STYLE.align);
  });

  it("lets a genuinely-default (unset) paragraph align pick up the band", () => {
    const { table } = headerTable(); // align === default, but no provenance recorded
    const rows = bakeTableStyleRows(table, alignBand, { Align: alignBand }, { firstRow: true }, {
      char: DEFAULT_CHAR_STYLE,
      para: DEFAULT_PARA_STYLE,
      explicitPara: new WeakMap(),
    });
    expect(bakedAlign(rows)).toBe("center");
  });
});

describe("cellCondFlags banding", () => {
  it("alternates band1/band2 across body rows, excluding the header", () => {
    const opts = { firstRow: true, bandRows: true };
    expect(cellCondFlags(0, 0, 5, 2, opts).rowBand).toBeUndefined(); // header
    expect(cellCondFlags(1, 0, 5, 2, opts).rowBand).toBe(1);
    expect(cellCondFlags(2, 0, 5, 2, opts).rowBand).toBe(2);
    expect(cellCondFlags(3, 0, 5, 2, opts).rowBand).toBe(1);
  });
});
