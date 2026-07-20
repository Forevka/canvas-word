// Single source for model → WordprocessingML value mappings shared across the
// exporter (documentXml, styleProps) and the headless TOC generator. Keeping
// these in one place prevents the maps from drifting apart (they already had).

import type { ParaStyle } from "@cw/shared";
import { pxToEighthPoints } from "../units";
import { el } from "./xmlWrite";

/** highlight hex → Word highlight color name (w:highlight w:val). Single-sourced
 *  in shared (model/highlight.ts) with the importer's name → hex map, so the two
 *  can't drift apart. */
export { HIGHLIGHT_NAME } from "@cw/shared";

/** paragraph align → w:jc w:val. */
export const JC: Record<ParaStyle["align"], string> = { left: "left", center: "center", right: "right", justify: "both" };

/** tab align → w:tab w:val. */
export const TAB_VAL: Record<string, string> = { left: "left", center: "center", right: "right", decimal: "decimal" };

/** tab leader → w:tab w:leader. */
export const TAB_LEADER: Record<string, string> = { dot: "dot", dash: "hyphen", underscore: "underscore" };

/** Normalize a model color to an OOXML color value: 6 lowercase hex digits (or the
 *  literal "auto"). Strips a leading "#", and — crucially — expands 3-digit CSS
 *  shorthand (#fff → ffffff), which Word's schema rejects as a color value. Every
 *  color attribute in the exporter (run w:color, w:shd w:fill, border w:color, page
 *  color) routes through here, so all of them stay schema-valid (issue #193). */
export const hexColor = (c: string): string => {
  if (c === "auto") return "auto";
  const h = c.replace(/^#/, "").toLowerCase();
  if (/^[0-9a-f]{3}$/.test(h)) return h.replace(/(.)/g, "$1$1"); // #abc → aabbcc
  return h;
};

/** A border edge in model terms — shared by cell, table, paragraph, and run
 *  borders (CellBorder and friends are all structurally this). */
export interface BorderEdgeSpec {
  color: string;
  widthPx: number;
  style?: string | undefined;
}

/** One border edge element (`w:top`, `w:bottom`, `w:bdr`, …): the ONE encoding of
 *  a model border edge as WordprocessingML, used by every border emitter in
 *  documentXml and styleProps. This exact mapping was previously copy-pasted five
 *  times and is precisely the kind of duplication this module exists to prevent. */
export function borderEdgeXml(name: string, spec: BorderEdgeSpec | undefined): string {
  if (!spec) return "";
  const val = spec.style === "double" ? "double" : spec.style === "dashed" ? "dashed" : spec.style === "dotted" ? "dotted" : "single";
  return el("w:" + name, { "w:val": val, "w:sz": pxToEighthPoints(spec.widthPx), "w:space": 0, "w:color": hexColor(spec.color) });
}
