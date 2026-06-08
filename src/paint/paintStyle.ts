// Shared paint styling — the constants and per-run styling DECISIONS that the
// canvas renderer (src/paint/renderer.ts) and the PDF painter
// (src/export/pdf/paintBlock.ts) must keep in lockstep. Centralising the magic
// numbers and the link/colour rules stops the two painters from silently
// drifting; the actual draw calls stay separate (canvas vs pdfkit are different
// APIs). Pure data + math — no DOM, no rendering backend.

import type { CharStyle, TabLeader } from "../model/document";

// --- colours ---------------------------------------------------------------
/** Unstyled/native table grid (when a cell carries no explicit borders). */
export const DEFAULT_GRID_COLOR = "#c0c4c9";
/** External hyperlinks paint in this blue (Word's affordance). */
export const EXTERNAL_LINK_COLOR = "#0b57d0";
/** In-document anchors that inherited a Hyperlink blue read as plain text. */
export const ANCHOR_TEXT_COLOR = "#202124";
export const FOOTNOTE_RULE_COLOR = "#80868b";
export const TOC_LEADER_COLOR = "#9aa0a6";
export const IMAGE_PLACEHOLDER_COLOR = "#f1f3f4";

/** Office "Hyperlink" character-style blues, normalised to text colour when they
 *  arrive on an in-document anchor (TOC/cross-ref) so those read as plain text. */
export const HYPERLINK_BLUES = new Set([
  "#0563c1", "#0000ff", "#0000ee", "#0b57d0", "#0066cc", "#1155cc",
]);

/** A leader/TOC colour that inherited a Hyperlink blue → text colour. */
export const normalizeLinkBlue = (color: string): string =>
  HYPERLINK_BLUES.has(color.toLowerCase()) ? ANCHOR_TEXT_COLOR : color;

// --- sub/superscript -------------------------------------------------------
/** Sub/superscript is MEASURED and PAINTED at this fraction of the run size
 *  (metrics.charStyleToFont scales the font; the painters scale to match). */
export const SUB_SUPER_SCALE = 0.65;
const SUPER_SHIFT = -0.38;
const SUB_SHIFT = 0.16;

/** Baseline shift (px) for a sub/superscript run — fraction of the ORIGINAL size. */
export function verticalShift(vertical: CharStyle["verticalAlign"], fontSizePx: number): number {
  if (vertical === "super") return SUPER_SHIFT * fontSizePx;
  if (vertical === "sub") return SUB_SHIFT * fontSizePx;
  return 0;
}

// --- text decorations ------------------------------------------------------
/** Underline / strikethrough thickness (px). */
export const decorationThickness = (fontSizePx: number): number => Math.max(1, fontSizePx / 14);
/** Underline distance below the baseline (the run's vShift is added by the caller). */
export const UNDERLINE_OFFSET_PX = 1.5;
/** Strikethrough offset relative to the baseline (above it). */
export const strikeOffset = (fontSizePx: number): number => -0.28 * fontSizePx;

/** The colour + decoration decision for a run, with link normalisation applied. */
export interface RunPaint {
  color: string;
  underline: boolean;
  strike: boolean;
  /** External (non-anchor) link — paints blue+underlined; the PDF also annotates it. */
  externalLink: boolean;
}
export function runPaint(style: CharStyle): RunPaint {
  const anchor = style.link !== undefined && style.link.startsWith("#");
  const externalLink = style.link !== undefined && !anchor;
  let color = style.color;
  if (externalLink) color = EXTERNAL_LINK_COLOR;
  else if (anchor && HYPERLINK_BLUES.has(style.color.toLowerCase())) color = ANCHOR_TEXT_COLOR;
  return {
    color,
    underline: externalLink || (!!style.underline && !anchor),
    strike: !!style.strikethrough,
    externalLink,
  };
}

// --- leaders (tab + TOC) ---------------------------------------------------
/** Tab-leader dash pattern; [] = solid (underscore/none). */
export function leaderDash(kind: TabLeader): number[] {
  return kind === "dot" ? [1, 3] : kind === "dash" ? [4, 3] : [];
}
export const leaderWidth = (fontSizePx: number): number => Math.max(1, fontSizePx / 14);
export const TOC_LEADER_DASH: number[] = [1, 4];
/** Gap (px) between an entry's text/number and its dot leader. */
export const TOC_LEADER_GAP_PX = 8;

// --- footnote rule ---------------------------------------------------------
/** The separator rule spans this fraction of the content width (Word style). */
export const FOOTNOTE_RULE_WIDTH_FRACTION = 1 / 3;

// --- cell borders ----------------------------------------------------------
/** Hairlines clamp up so a thin rule stays visible. */
export const cellBorderWidth = (widthPx: number): number => Math.max(0.5, widthPx);
/** Dash pattern for a cell edge of stroke width w; [] = solid. */
export function cellBorderDash(style: string | undefined, w: number): number[] {
  return style === "dashed" ? [w * 3, w * 2] : style === "dotted" ? [w, w * 1.5] : [];
}
/** Offset of the inner stroke of a "double" border. */
export const doubleBorderGap = (w: number): number => w + 1;
