// Single source for model → WordprocessingML value mappings shared across the
// exporter (documentXml, styleProps) and the headless TOC generator. Keeping
// these in one place prevents the maps from drifting apart (they already had).

import type { ParaStyle } from "@cw/shared";

/** highlight hex → Word highlight color name (w:highlight w:val). */
export const HIGHLIGHT_NAME: Record<string, string> = {
  "#ffff00": "yellow", "#00ff00": "green", "#00ffff": "cyan", "#ff00ff": "magenta",
  "#0000ff": "blue", "#ff0000": "red", "#000080": "darkBlue", "#008080": "darkCyan",
  "#008000": "darkGreen", "#800080": "darkMagenta", "#800000": "darkRed", "#808000": "darkYellow",
  "#808080": "darkGray", "#c0c0c0": "lightGray", "#000000": "black", "#ffffff": "white",
};

/** paragraph align → w:jc w:val. */
export const JC: Record<ParaStyle["align"], string> = { left: "left", center: "center", right: "right", justify: "both" };

/** tab align → w:tab w:val. */
export const TAB_VAL: Record<string, string> = { left: "left", center: "center", right: "right", decimal: "decimal" };

/** tab leader → w:tab w:leader. */
export const TAB_LEADER: Record<string, string> = { dot: "dot", dash: "hyphen", underscore: "underscore" };

/** strip a leading "#" and lowercase a hex color for OOXML w:val attributes. */
export const hexColor = (c: string): string => c.replace(/^#/, "").toLowerCase();
