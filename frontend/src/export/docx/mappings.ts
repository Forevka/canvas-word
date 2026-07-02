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

/** strip a leading "#" and lowercase a hex color for OOXML w:val attributes. */
export const hexColor = (c: string): string => c.replace(/^#/, "").toLowerCase();

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
