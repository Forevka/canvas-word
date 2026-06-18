import { describe, expect, it } from "vitest";
import { cellCondFlags, effectiveCellProps, resolveTableStyle, type TableStyle } from "./tableStyles";

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

describe("cellCondFlags banding", () => {
  it("alternates band1/band2 across body rows, excluding the header", () => {
    const opts = { firstRow: true, bandRows: true };
    expect(cellCondFlags(0, 0, 5, 2, opts).rowBand).toBeUndefined(); // header
    expect(cellCondFlags(1, 0, 5, 2, opts).rowBand).toBe(1);
    expect(cellCondFlags(2, 0, 5, 2, opts).rowBand).toBe(2);
    expect(cellCondFlags(3, 0, 5, 2, opts).rowBand).toBe(1);
  });
});
