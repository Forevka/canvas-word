// Table border & shading decode + cascade resolution.
//
// OOXML resolves a cell edge through: table style (styles.xml) → table direct
// (w:tblBorders) → cell direct (w:tcBorders), with the table layer choosing its
// OUTER edge (top/left/bottom/right) on the table boundary and its INTERIOR
// edge (insideH/insideV) between cells. A cell's own w:tcBorders are already
// written per-edge by Word, so they map straight onto the named edge.
//
// "First definition wins" (incl. an explicit w:val="nil" → no line), so the
// cascade picks the most-specific source that DEFINES the edge, even to suppress.

import type { CellBorder, CellBorders } from "@cw/shared";
import type { IRBorders, IRRawBorder } from "./types";
import { round2 } from "./units";
import { attr, el, numAttr, type XmlNode } from "./xml";

const EDGES: [keyof IRBorders, string][] = [
  ["top", "w:top"],
  ["left", "w:left"],
  ["bottom", "w:bottom"],
  ["right", "w:right"],
  ["insideH", "w:insideH"],
  ["insideV", "w:insideV"],
];

export function decodeBorders(bordersEl: XmlNode | undefined): IRBorders | undefined {
  if (!bordersEl) return undefined;
  const out: IRBorders = {};
  for (const [key, tag] of EDGES) {
    const e = el(bordersEl, tag);
    if (!e) continue;
    const raw: IRRawBorder = { val: attr(e, "w:val") ?? "single" };
    const sz = numAttr(e, "w:sz");
    if (sz !== undefined) raw.sizeEighthPt = sz;
    const color = attr(e, "w:color");
    if (color) raw.color = color;
    out[key] = raw;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

/** w:shd → CSS fill, or undefined when there's no real fill. */
export function decodeShdFill(shd: XmlNode | undefined): string | undefined {
  if (!shd) return undefined;
  const fill = attr(shd, "w:fill");
  if (fill && fill !== "auto") return `#${fill.toLowerCase()}`;
  // Pattern shading with a foreground color but "auto" fill — approximate with
  // the pattern color so e.g. a solid-color cell still shows.
  const val = attr(shd, "w:val");
  const color = attr(shd, "w:color");
  if (val && val !== "clear" && val !== "nil" && color && color !== "auto") {
    return `#${color.toLowerCase()}`;
  }
  return undefined;
}

/** Merge two border sets edge-by-edge (child overrides parent) — for resolving
 *  a table style's basedOn chain. */
export function mergeBorders(base: IRBorders | undefined, over: IRBorders | undefined): IRBorders | undefined {
  if (!base) return over;
  if (!over) return base;
  return { ...base, ...over };
}

export interface BorderSources {
  cell?: IRBorders | undefined;
  table?: IRBorders | undefined;
  style?: IRBorders | undefined;
}

export interface CellPosition {
  top: boolean;
  bottom: boolean;
  left: boolean;
  right: boolean;
}

/** Collapse the cascade for one cell into concrete per-edge borders. */
export function resolveCellBorders(src: BorderSources, pos: CellPosition): CellBorders {
  const pick = (named: keyof IRBorders, inside: keyof IRBorders, outer: boolean): IRRawBorder | undefined =>
    // Cell edges are named directly; the table/style layers use the named edge
    // on the boundary and the inside edge between cells.
    src.cell?.[named] ?? src.table?.[outer ? named : inside] ?? src.style?.[outer ? named : inside];

  const out: CellBorders = {};
  const top = toCellBorder(pick("top", "insideH", pos.top));
  if (top) out.top = top;
  const bottom = toCellBorder(pick("bottom", "insideH", pos.bottom));
  if (bottom) out.bottom = bottom;
  const left = toCellBorder(pick("left", "insideV", pos.left));
  if (left) out.left = left;
  const right = toCellBorder(pick("right", "insideV", pos.right));
  if (right) out.right = right;
  return out;
}

/** Map an IRBorders' four outer edges to model CellBorders (used by table-style
 *  conditional bands, which apply per-cell). Inside edges (insideH/V) are not
 *  represented in per-cell CellBorders and are dropped. */
export function cellBordersFromIR(b: IRBorders | undefined): CellBorders | undefined {
  if (!b) return undefined;
  const out: CellBorders = {};
  const t = toCellBorder(b.top); if (t) out.top = t;
  const r = toCellBorder(b.right); if (r) out.right = r;
  const bot = toCellBorder(b.bottom); if (bot) out.bottom = bot;
  const l = toCellBorder(b.left); if (l) out.left = l;
  return Object.keys(out).length > 0 ? out : undefined;
}

function toCellBorder(raw: IRRawBorder | undefined): CellBorder | undefined {
  if (!raw || raw.val === "nil" || raw.val === "none") return undefined;
  // w:sz is eighths of a point: px = sz/8 pt × 96/72 = sz/6. Hairlines clamp up
  // so a 2-eighths (0.25pt) rule stays visible at 96dpi.
  const widthPx = round2(Math.max(0.5, (raw.sizeEighthPt ?? 4) / 6));
  const color = !raw.color || raw.color === "auto" ? "#000000" : `#${raw.color.toLowerCase()}`;
  const border: CellBorder = { color, widthPx };
  const style = mapBorderStyle(raw.val);
  if (style !== "single") border.style = style;
  return border;
}

function mapBorderStyle(val: string): NonNullable<CellBorder["style"]> {
  if (val === "double") return "double";
  if (/dash/i.test(val)) return "dashed";
  if (/dot/i.test(val)) return "dotted";
  return "single"; // single, thick, wave, 3D variants → a plain line
}
