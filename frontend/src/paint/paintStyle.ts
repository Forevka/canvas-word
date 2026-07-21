// Shared paint styling — the constants and per-run styling DECISIONS that the
// canvas renderer (src/paint/renderer.ts) and the PDF painter
// (src/export/pdf/paintBlock.ts) must keep in lockstep. Centralising the magic
// numbers and the link/colour rules stops the two painters from silently
// drifting; the actual draw calls stay separate (canvas vs pdfkit are different
// APIs). Pure data + math — no DOM, no rendering backend.

import type { CharStyle, PageBorderEdge, PageBorders, ShapeDash, TabLeader, UnderlineStyle } from "@cw/shared";
import { ptToPx } from "@cw/shared";
import type { PlacedShape } from "../layout/layoutTree";

// --- colours ---------------------------------------------------------------
/** Unstyled/native table grid (when a cell carries no explicit borders). */
export const DEFAULT_GRID_COLOR = "#c0c4c9";
/** External hyperlinks paint in this blue (Word's affordance). */
export const EXTERNAL_LINK_COLOR = "#0b57d0";
/** UI-accent blue for editor chrome affordances (e.g. the band-edit boundary). */
export const ACCENT_BLUE = "#1a73e8";
/** In-document anchors that inherited a Hyperlink blue read as plain text. */
export const ANCHOR_TEXT_COLOR = "#202124";
export const FOOTNOTE_RULE_COLOR = "#80868b";
/** Non-printing formatting marks (space dots, tab arrows, pilcrows, line breaks)
 *  drawn only when "show formatting marks" is on. A muted gray-blue, like Word. */
export const FORMATTING_MARK_COLOR = "#7a8aa0";
export const TOC_LEADER_COLOR = "#9aa0a6";
export const IMAGE_PLACEHOLDER_COLOR = "#f1f3f4";

// --- drawing shapes --------------------------------------------------------
/** Theme-neutral defaults for a shape whose fill/stroke are absent (an author
 *  can still request an explicit no-fill / no-outline via `{ none: true }`). */
export const SHAPE_DEFAULT_FILL = "#d9e2f0";
export const SHAPE_DEFAULT_STROKE_COLOR = "#41719c";
export const SHAPE_DEFAULT_STROKE_WIDTH_PT = 1;

/** The on/off px pattern for an outline dash preset, scaled by the stroke width so
 *  a thicker line gets proportionally longer dashes (matching Word/DrawingML). An
 *  empty array (solid, or absent) means "no dash" — the painters skip setLineDash.
 *  Shared by the canvas renderer and the PDF painter so a dashed outline is
 *  pixel-identical on screen and in the export. */
export function shapeDashArray(dash: ShapeDash | undefined, widthPx: number): number[] {
  const w = Math.max(1, widthPx);
  switch (dash) {
    case "dot": return [w, 2 * w];
    case "dash": return [4 * w, 3 * w];
    case "lgDash": return [8 * w, 3 * w];
    case "dashDot": return [4 * w, 3 * w, w, 3 * w];
    default: return []; // "solid" / undefined
  }
}

/** Resolve a placed shape's fill + stroke into concrete paint values the canvas and
 *  PDF painters both consume (keeping them in lockstep). `fill`/`stroke` come back
 *  null when there is nothing to paint (explicit `none`, or a line's absent fill).
 *  A resolved stroke carries its dash as a ready-to-use px pattern (empty = solid). */
export function resolveShapePaint(shape: PlacedShape): {
  fill: string | null;
  stroke: { color: string; widthPx: number; dash: number[] } | null;
} {
  // Fill: explicit color, explicit none → null, absent → default (but a preset line
  // is a stroke-only geometry, so it never fills; a custom path always may fill).
  let fill: string | null;
  if (shape.fill && "none" in shape.fill) fill = null;
  else if (shape.fill && "color" in shape.fill) fill = shape.fill.color;
  else fill = shape.preset === "line" && !shape.custom ? null : SHAPE_DEFAULT_FILL;
  // Stroke: explicit (unless none), absent → default.
  let stroke: { color: string; widthPx: number; dash: number[] } | null;
  if (shape.stroke && "none" in shape.stroke) {
    stroke = null;
  } else if (shape.stroke) {
    const widthPx = Math.max(0.5, ptToPx(shape.stroke.widthPt));
    stroke = { color: shape.stroke.color, widthPx, dash: shapeDashArray(shape.stroke.dash, widthPx) };
  } else {
    stroke = { color: SHAPE_DEFAULT_STROKE_COLOR, widthPx: ptToPx(SHAPE_DEFAULT_STROKE_WIDTH_PT), dash: [] };
  }
  return { fill, stroke };
}

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

/** Total baseline shift (px, DOWN = positive) a run paints with: the sub/superscript
 *  shift PLUS the explicit w:position raise/lower (positionPx>0 raises ⇒ shifts up,
 *  i.e. subtracts). Both painters add this to the baseline. */
export function runVerticalShift(style: CharStyle): number {
  return verticalShift(style.verticalAlign, style.fontSizePx) - (style.positionPx ?? 0);
}

/** Horizontal glyph scale factor for a run (OOXML w:w). 1 = no scaling. The
 *  layout reserves the scaled advance; the painters stretch the glyphs to match. */
export function widthScale(style: CharStyle): number {
  const pct = style.widthScalePct;
  return pct !== undefined && pct > 0 ? pct / 100 : 1;
}

/** The two rule offsets (px, relative to the baseline) of a double strikethrough
 *  (OOXML w:dstrike) — straddling the single-strike line so both rules stay on the
 *  glyph body. Shared by the canvas + PDF painters. */
export function doubleStrikeOffsets(fontSizePx: number): [number, number] {
  const base = strikeOffset(fontSizePx);
  const gap = Math.max(1.5, fontSizePx / 11);
  return [base - gap / 2, base + gap / 2];
}
// --- small caps ------------------------------------------------------------
/** Small-caps (w:smallCaps): originally-lowercase letters are uppercased and drawn
 *  at this fraction of the run size. The layout BAKES this scale into a reduced
 *  fontSizePx on the affected sub-runs, so the painters need no small-caps logic —
 *  they just draw the (already uppercased) glyphs at the (already reduced) size.
 *  ~0.78 approximates Word's small-cap height across common fonts. */
export const SMALL_CAPS_SCALE = 0.78;

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
  /** Underline line style — "single" (and link underlines) paint as a plain rule. */
  underlineStyle: UnderlineStyle;
  /** Colour of the underline rule: the run's explicit underlineColor when set,
   *  else the (possibly link-normalised) text colour. */
  underlineColor: string;
  strike: boolean;
  /** External (non-anchor) link — paints blue+underlined; the PDF also annotates it. */
  externalLink: boolean;
}
export function runPaint(style: CharStyle, linkColor: string = EXTERNAL_LINK_COLOR): RunPaint {
  const anchor = style.link !== undefined && style.link.startsWith("#");
  const externalLink = style.link !== undefined && !anchor;
  let color = style.color;
  if (externalLink) color = linkColor;
  else if (anchor && HYPERLINK_BLUES.has(style.color.toLowerCase())) color = ANCHOR_TEXT_COLOR;
  // A link affordance always draws a plain solid rule in the link colour; an
  // explicit underline keeps its style and (optional) colour.
  const underline = externalLink || (!!style.underline && !anchor);
  return {
    color,
    underline,
    underlineStyle: externalLink ? "single" : (style.underline ? (style.underlineStyle ?? "single") : "single"),
    underlineColor: !externalLink && style.underline && style.underlineColor ? style.underlineColor : color,
    strike: !!style.strikethrough,
    externalLink,
  };
}

// --- list bullet glyphs ----------------------------------------------------
// The three default bullet levels (shared/src/model/lists.ts BULLETS) are drawn
// as VECTOR SHAPES rather than text in both painters: ◦ (U+25E6 WHITE BULLET) and
// ▪ (U+25AA BLACK SMALL SQUARE) are absent from the bundled PDF Latin font subset,
// so painting them as glyphs yields .notdef/tofu in PDF export (issue #99). A
// shared shape keeps the on-screen canvas and the PDF pixel-identical and removes
// any dependence on font glyph coverage. Non-bullet markers (numbers, custom
// chars) still paint as text via the normal marker path.

/** U+2022 BULLET — a filled disc. */
export const BULLET_DISC = "•";
/** U+25E6 WHITE BULLET — a hollow ring. */
export const BULLET_RING = "◦";
/** U+25AA BLACK SMALL SQUARE — a filled square. */
export const BULLET_SQUARE = "▪";

// Geometry as a fraction of the marker font size. Tuned to sit close to where the
// glyph's ink would land: centred a touch right of the marker's left edge and a
// little above the baseline, at roughly a glyph-sized diameter.
const BULLET_CENTER_X_EM = 0.22;
const BULLET_CENTER_Y_EM = 0.31;
const DISC_RADIUS_EM = 0.135;
const RING_RADIUS_EM = 0.13;
const RING_STROKE_EM = 0.05;
const SQUARE_SIZE_EM = 0.22;

/** A backend-agnostic description of a default-bullet marker shape, shared by the
 *  canvas renderer and the PDF painter so both draw an identical bullet. */
export type BulletShape =
  | { kind: "disc"; cx: number; cy: number; r: number }
  | { kind: "ring"; cx: number; cy: number; r: number; lineWidth: number }
  | { kind: "square"; x: number; y: number; size: number };

/** The vector shape for a default-bullet marker, or `undefined` for any other
 *  marker text (numbers, custom glyphs) which the painters render as text. `x` is
 *  the marker's left edge and `baselineY` its text baseline — the same coordinates
 *  the glyph would have been painted at. */
export function bulletShapeFor(
  text: string,
  fontSizePx: number,
  x: number,
  baselineY: number,
): BulletShape | undefined {
  const cx = x + fontSizePx * BULLET_CENTER_X_EM;
  const cy = baselineY - fontSizePx * BULLET_CENTER_Y_EM;
  switch (text) {
    case BULLET_DISC:
      return { kind: "disc", cx, cy, r: fontSizePx * DISC_RADIUS_EM };
    case BULLET_RING:
      return { kind: "ring", cx, cy, r: fontSizePx * RING_RADIUS_EM, lineWidth: fontSizePx * RING_STROKE_EM };
    case BULLET_SQUARE: {
      const size = fontSizePx * SQUARE_SIZE_EM;
      return { kind: "square", x: cx - size / 2, y: cy - size / 2, size };
    }
    default:
      return undefined;
  }
}

/** How an underline of a given style is stroked, shared by the canvas + PDF
 *  painters so both reproduce double/dotted/dashed/wave identically. */
export interface UnderlinePlan {
  /** Line-dash pattern (canvas setLineDash / pdfkit .dash); [] = solid. */
  dash: number[];
  /** Stroke two parallel lines (the "double" style). */
  double: boolean;
  /** Stroke a sine wave instead of a straight line (the "wave" style). */
  wave: boolean;
  /** Stroke thickness (px). */
  thickness: number;
}

/** Resolve an underline style + font size to its concrete stroke plan. */
export function underlinePlan(style: UnderlineStyle, fontSizePx: number): UnderlinePlan {
  const w = decorationThickness(fontSizePx);
  switch (style) {
    case "double": return { dash: [], double: true, wave: false, thickness: w };
    case "thick": return { dash: [], double: false, wave: false, thickness: Math.max(2, w * 2) };
    case "dotted": return { dash: [w, w * 1.5], double: false, wave: false, thickness: w };
    case "dash": return { dash: [w * 3, w * 2], double: false, wave: false, thickness: w };
    case "dotDash": return { dash: [w * 3, w * 2, w, w * 2], double: false, wave: false, thickness: w };
    case "dotDotDash": return { dash: [w * 3, w * 2, w, w * 2, w, w * 2], double: false, wave: false, thickness: w };
    case "wave": return { dash: [], double: false, wave: true, thickness: w };
    default: return { dash: [], double: false, wave: false, thickness: w }; // single
  }
}

/** Gap (px) between the two rules of a "double" underline. */
export const doubleUnderlineGap = (thickness: number): number => thickness + 1;

/** Polyline samples of a wavy underline over [x, x+width], centred on `yCenter`.
 *  Both painters stroke through these points so the wave matches pixel-for-pixel. */
export function underlineWavePoints(x: number, yCenter: number, width: number, fontSizePx: number): { x: number; y: number }[] {
  const amp = Math.max(0.75, fontSizePx / 18);
  const period = Math.max(4, fontSizePx / 3);
  const steps = Math.max(2, Math.ceil((width / period) * 4));
  const pts: { x: number; y: number }[] = [];
  for (let i = 0; i <= steps; i++) {
    const px = x + (width * i) / steps;
    pts.push({ x: px, y: yCenter + amp * Math.sin((2 * Math.PI * (px - x)) / period) });
  }
  return pts;
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

// --- column separators -----------------------------------------------------
/** Thin gray rule drawn between newspaper columns (w:cols/@w:sep). */
export const COLUMN_SEPARATOR_COLOR = "#c0c4c9";

// --- paragraph borders (w:pBdr) --------------------------------------------
/** Border-to-text padding (px) for a paragraph border box (w:pBdr) when no
 *  w:space is imported. Word's UI default is 1pt but reads cramped on screen;
 *  ~4px keeps a wide/double rule clear of the glyphs. The decor box is expanded
 *  OUTWARD by this on every side, so the border sits outside the text and the
 *  shading fill reaches the border. */
export const PARA_BORDER_PAD_PX = 4;

/** The padded paragraph-decoration box (shading fill + border rectangle), in
 *  page (px) coords. `d.pad` (0 when shading-only / undecorated-border) expands
 *  the text box outward symmetrically. Shared so the canvas renderer and the PDF
 *  painter draw the border and fill at exactly the same place. */
export function paraDecorBox(
  x: number,
  y: number,
  d: { width: number; height: number; pad?: number },
): { x: number; y: number; width: number; height: number } {
  const p = d.pad ?? 0;
  return { x: x - p, y: y - p, width: d.width + 2 * p, height: d.height + 2 * p };
}

// --- page borders (w:pgBorders) --------------------------------------------
/** Default offset (px) of a page border from the page edge / text when an edge
 *  omits w:space. ~24pt is Word's typical page-border offset. */
const PAGE_BORDER_DEFAULT_SPACE = 24;

/** One stroked line segment of a page border, in page (px) coords. */
export interface PageBorderSegment {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  widthPx: number;
  color: string;
  /** [] = solid. */
  dash: number[];
}

/** Geometry of a page border box → the line segments to stroke. Shared so the
 *  canvas renderer and the PDF painter draw identical borders. "double" yields
 *  two parallel segments; "none"/absent edges are skipped. */
export function pageBorderSegments(
  borders: PageBorders,
  pageW: number,
  pageH: number,
  margin: { top: number; right: number; bottom: number; left: number },
): PageBorderSegment[] {
  const fromText = borders.offsetFrom === "text";
  const sp = (e?: PageBorderEdge): number => e?.spacePx ?? PAGE_BORDER_DEFAULT_SPACE;
  // Each edge's distance, then the box bounds (per-edge space; usually uniform).
  const xL = fromText ? margin.left - sp(borders.left) : sp(borders.left);
  const xR = fromText ? pageW - margin.right + sp(borders.right) : pageW - sp(borders.right);
  const yT = fromText ? margin.top - sp(borders.top) : sp(borders.top);
  const yB = fromText ? pageH - margin.bottom + sp(borders.bottom) : pageH - sp(borders.bottom);
  const out: PageBorderSegment[] = [];
  const push = (e: PageBorderEdge | undefined, x1: number, y1: number, x2: number, y2: number): void => {
    if (!e || e.style === "none") return;
    const w = e.style === "thick" ? Math.max(2, e.widthPx) : Math.max(0.5, e.widthPx);
    const color = e.color === "auto" ? "#000000" : e.color;
    const dash = e.style === "dashed" ? [w * 3, w * 2] : e.style === "dotted" ? [w, w * 1.5] : [];
    out.push({ x1, y1, x2, y2, widthPx: w, color, dash });
    if (e.style === "double") {
      const g = w + 1; // inset the second line toward the box interior
      const horizontal = y1 === y2;
      if (horizontal) {
        const dy = y1 < pageH / 2 ? g : -g;
        out.push({ x1, y1: y1 + dy, x2, y2: y2 + dy, widthPx: w, color, dash: [] });
      } else {
        const dx = x1 < pageW / 2 ? g : -g;
        out.push({ x1: x1 + dx, y1, x2: x2 + dx, y2, widthPx: w, color, dash: [] });
      }
    }
  };
  push(borders.top, xL, yT, xR, yT);
  push(borders.bottom, xL, yB, xR, yB);
  push(borders.left, xL, yT, xL, yB);
  push(borders.right, xR, yT, xR, yB);
  return out;
}

// --- cell borders ----------------------------------------------------------
/** Hairlines clamp up so a thin rule stays visible. */
export const cellBorderWidth = (widthPx: number): number => Math.max(0.5, widthPx);
/** Dash pattern for a cell edge of stroke width w; [] = solid. */
export function cellBorderDash(style: string | undefined, w: number): number[] {
  return style === "dashed" ? [w * 3, w * 2] : style === "dotted" ? [w, w * 1.5] : [];
}
/** Offset of the inner stroke of a "double" border. */
export const doubleBorderGap = (w: number): number => w + 1;
