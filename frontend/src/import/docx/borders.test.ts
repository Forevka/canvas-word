// Table border decode + cascade. decodeBorders/decodeShdFill take parsed XML
// nodes; the cascade (resolveCellBorders) and IR->model mappers take IRBorders
// objects. We assert the documented "first definition wins" cascade, the
// outer-vs-inside edge selection, the eighth-point->px width math, the hairline
// clamp, and the val->style mapping.

import { describe, expect, it } from "vitest";
import {
  cellBordersFromIR,
  decodeBorders,
  decodeShdFill,
  mergeBorders,
  resolveCellBorders,
} from "./borders";
import { parseXml } from "./xml";
import type { IRBorders } from "./types";

// Parse a w:tblBorders/w:tcBorders fragment into the single XmlNode the decoders want.
const node = (xml: string) => parseXml(`<w:x xmlns:w="urn">${xml}</w:x>`, "t")[0]!;
// Parse a single self-contained element (e.g. w:shd) as the root node.
const elem = (xml: string) => parseXml(`<root xmlns:w="urn">${xml}</root>`, "t")[0]!.children[0] as never;

const ALL_POS = { top: true, bottom: true, left: true, right: true };
const NO_POS = { top: false, bottom: false, left: false, right: false };

describe("decodeBorders", () => {
  it("returns undefined for a missing element", () => {
    expect(decodeBorders(undefined)).toBeUndefined();
  });

  it("decodes per-edge val/size/color", () => {
    const b = decodeBorders(
      node(`<w:top w:val="single" w:sz="8" w:color="FF0000"/>` + `<w:left w:val="double" w:sz="12"/>`),
    );
    expect(b).toEqual({
      top: { val: "single", sizeEighthPt: 8, color: "FF0000" },
      left: { val: "double", sizeEighthPt: 12 },
    });
  });

  it("defaults a missing w:val to 'single'", () => {
    const b = decodeBorders(node(`<w:top w:sz="4"/>`));
    expect(b!.top).toEqual({ val: "single", sizeEighthPt: 4 });
  });

  it("decodes inside edges (insideH/insideV)", () => {
    const b = decodeBorders(node(`<w:insideH w:val="single"/><w:insideV w:val="single"/>`));
    expect(Object.keys(b!)).toEqual(["insideH", "insideV"]);
  });

  it("returns undefined when the element has no recognized edges", () => {
    expect(decodeBorders(node(`<w:bogus/>`))).toBeUndefined();
  });
});

describe("decodeShdFill", () => {
  it("returns undefined for missing shd", () => {
    expect(decodeShdFill(undefined)).toBeUndefined();
  });

  it("maps a concrete w:fill to a lowercased CSS hex", () => {
    expect(decodeShdFill(elem(`<w:shd w:fill="FFFF00"/>`))).toBe("#ffff00");
  });

  it("treats w:fill='auto' as no fill", () => {
    expect(decodeShdFill(elem(`<w:shd w:fill="auto"/>`))).toBeUndefined();
  });

  it("approximates a pattern fill with the foreground color", () => {
    expect(decodeShdFill(elem(`<w:shd w:val="pct50" w:fill="auto" w:color="00FF00"/>`))).toBe("#00ff00");
  });

  it("returns undefined for clear/nil patterns or auto color", () => {
    expect(decodeShdFill(elem(`<w:shd w:val="clear" w:fill="auto" w:color="00FF00"/>`))).toBeUndefined();
    expect(decodeShdFill(elem(`<w:shd w:val="pct50" w:fill="auto" w:color="auto"/>`))).toBeUndefined();
  });
});

describe("mergeBorders", () => {
  const a: IRBorders = { top: { val: "single" }, left: { val: "single" } };
  const b: IRBorders = { top: { val: "double" }, right: { val: "single" } };

  it("returns the other side when one is undefined", () => {
    expect(mergeBorders(undefined, b)).toBe(b);
    expect(mergeBorders(a, undefined)).toBe(a);
  });

  it("lets the override win edge-by-edge", () => {
    expect(mergeBorders(a, b)).toEqual({
      top: { val: "double" }, // overridden
      left: { val: "single" }, // from base
      right: { val: "single" }, // from override
    });
  });
});

describe("resolveCellBorders cascade", () => {
  it("prefers the cell layer over table and style (first definition wins)", () => {
    const out = resolveCellBorders(
      {
        cell: { top: { val: "single", sizeEighthPt: 24 } },
        table: { top: { val: "single", sizeEighthPt: 8 } },
        style: { top: { val: "single", sizeEighthPt: 4 } },
      },
      ALL_POS,
    );
    expect(out.top).toEqual({ color: "#000000", widthPx: 4 }); // 24/6 = 4px
  });

  it("uses the table's NAMED edge on a boundary cell", () => {
    const out = resolveCellBorders(
      { table: { top: { val: "single", sizeEighthPt: 12 }, insideH: { val: "single", sizeEighthPt: 6 } } },
      { ...ALL_POS, top: true },
    );
    expect(out.top).toEqual({ color: "#000000", widthPx: 2 }); // named top: 12/6 = 2
  });

  it("uses the table's INSIDE edge for an interior (non-boundary) edge", () => {
    const out = resolveCellBorders(
      { table: { top: { val: "single", sizeEighthPt: 12 }, insideH: { val: "single", sizeEighthPt: 6 } } },
      { ...NO_POS }, // not on any boundary -> top uses insideH
    );
    expect(out.top).toEqual({ color: "#000000", widthPx: 1 }); // insideH: 6/6 = 1
  });

  it("suppresses an edge whose most-specific source defines w:val='nil'", () => {
    const out = resolveCellBorders(
      { cell: { top: { val: "nil" } }, table: { top: { val: "single", sizeEighthPt: 8 } } },
      ALL_POS,
    );
    expect(out.top).toBeUndefined();
  });

  it("returns no edges for empty sources", () => {
    expect(resolveCellBorders({}, ALL_POS)).toEqual({});
  });
});

describe("toCellBorder math (via cellBordersFromIR)", () => {
  it("converts eighth-points to px with the /6 factor", () => {
    expect(cellBordersFromIR({ top: { val: "single", sizeEighthPt: 48 } })!.top).toEqual({
      color: "#000000",
      widthPx: 8,
    });
  });

  it("clamps hairlines up to 0.5px", () => {
    // 2 eighths = 0.333px -> clamped to 0.5
    expect(cellBordersFromIR({ top: { val: "single", sizeEighthPt: 2 } })!.top!.widthPx).toBe(0.5);
  });

  it("defaults a missing size to 4 eighths (0.67px)", () => {
    expect(cellBordersFromIR({ top: { val: "single" } })!.top!.widthPx).toBe(0.67);
  });

  it("defaults a missing/auto color to black", () => {
    expect(cellBordersFromIR({ top: { val: "single", color: "auto" } })!.top!.color).toBe("#000000");
    expect(cellBordersFromIR({ top: { val: "single", color: "AABBCC" } })!.top!.color).toBe("#aabbcc");
  });

  it("maps border val to a style (omitting 'single')", () => {
    expect(cellBordersFromIR({ top: { val: "double" } })!.top!.style).toBe("double");
    expect(cellBordersFromIR({ top: { val: "dashSmallGap" } })!.top!.style).toBe("dashed");
    expect(cellBordersFromIR({ top: { val: "dotted" } })!.top!.style).toBe("dotted");
    expect(cellBordersFromIR({ top: { val: "thick" } })!.top!.style).toBeUndefined(); // plain single
  });

  it("drops nil/none edges and inside edges, returning undefined when nothing remains", () => {
    expect(cellBordersFromIR({ top: { val: "nil" } })).toBeUndefined();
    expect(cellBordersFromIR({ insideH: { val: "single" } })).toBeUndefined();
    expect(cellBordersFromIR(undefined)).toBeUndefined();
  });
});
