// Layer 2 output: the LayoutTree — absolutely-positioned geometry the paint and
// input layers consume. Coordinates are CSS px, page-relative.

import type { BandContainer, CellBorders, CharStyle, PageBorders, TabLeader } from "@cw/shared";
import type { MathBox } from "./math/mathBox";

/** A same-styled slice of text placed on a line. One ctx.fillText call each.
 *  Per-cluster advances for caret math are computed lazily by geometry.ts and
 *  cached per fragment — most fragments are never hit-tested. */
export interface InlineFragment {
  blockId: string;
  startOffset: number; // UTF-16 offsets into the paragraph text
  endOffset: number;
  text: string;
  style: CharStyle;
  x: number;
  width: number;
  /** Justification: extra px per U+0020 (painted via ctx.wordSpacing, measured
   *  identically in geometry). Absent on non-justified lines. */
  wordSpacingPx?: number;
  /** Present only when pretext collapsed whitespace inside this fragment, making
   *  rendered text shorter than its model range. offsetMap[localIndex] = model
   *  offset delta from startOffset (length text.length + 1). Absent = identity. */
  offsetMap?: number[];
  /** Bidi embedding level (UAX#9). Even = left-to-right, odd = right-to-left.
   *  Absent (treated as 0) on the LTR fast path. Geometry reverses per-cluster
   *  caret math when this is odd, and paint relies on the canvas to shape/render
   *  the (single-level) fragment text in its own direction. */
  level?: number;
  /** Inline equation (OOXML inline m:oMath): this fragment is a single U+FFFC that
   *  paints as the typeset math box instead of text. `width` is the box width; the
   *  box glyphs/rules are relative to the line baseline. */
  equation?: MathBox;
}

export interface LineBox {
  y: number; // top, relative to the placed block
  height: number;
  ascent: number;
  fragments: InlineFragment[];
  /** Caret offset a fragment-less line represents (empty paragraph: 0; empty
   *  soft-break segment: the segment's start offset). Geometry indexes it. */
  emptyOffset?: number;
  /** Tab leaders (dot/dash/underscore fills) to paint in the gap a tab opened.
   *  x is block-relative (alignment already applied); paint draws on the baseline. */
  leaders?: { x1: number; x2: number; kind: TabLeader; color: string; fontSizePx: number }[];
  // ---- formatting-marks overlay (paint-only when the toggle is on) ----------
  /** Tab gaps on this line (block-relative, alignment applied). The marks overlay
   *  draws a "→" arrow in each; independent of `leaders` (a tab can have both). */
  tabArrows?: { x1: number; x2: number }[];
  /** This line is the FINAL line of its paragraph → draw a pilcrow (¶) after it. */
  paragraphEnd?: boolean;
  /** This line ends with a manual line break (soft return) → draw a "↵" after it. */
  lineBreak?: boolean;
}

export interface PlacedImage {
  src: string;
  width: number;
  height: number;
  /** Clip rect (block-absolute) for object-fit:cover — a sole image filling a
   *  tall cell is scaled to cover and clipped to the cell box. Absent = no clip. */
  clip?: { x: number; y: number; width: number; height: number };
  /** Behind-text anchored image (ImageBlock.anchor.behind): painted under the
   *  text and ignored by hit-testing so the text beneath stays selectable. */
  behind?: boolean;
  /** In-front-of-text anchored image (anchored, not behind): painted ABOVE the
   *  text layer (the engine's z-order pass moves it last in the page block list). */
  front?: boolean;
  /** Stacking order within the behind/front layer (ImageBlock.anchor.z). */
  z?: number;
}

/** A placed display equation — the typeset math box plus where its baseline sits
 *  within the block. Painters draw `box.glyphs`/`box.rules` at (blockX + glyph.x,
 *  blockY + baseline + glyph.y). */
export interface PlacedEquation {
  box: MathBox;
  width: number;
  height: number;
  /** Baseline offset from the placed block's top (= box.ascent). */
  baseline: number;
}

export interface PlacedTableCell {
  x: number; // absolute page coords for everything in tables
  y: number;
  width: number;
  height: number;
  /** Grid origin of this (possibly merged) cell — its top row and left column in
   *  the table grid. Hit-testing returns these so a drag stays anchored to one
   *  merged cell instead of flipping columns/rows across its colSpan/rowSpan. */
  originRow: number;
  originCol: number;
  /** Cell paragraphs as regular PlacedBlocks — geometry indexes them, so
   *  click/caret/selection/typing work inside cells with zero special cases. */
  blocks: PlacedBlock[];
  /** Resolved fill (CSS color) carried from the model cell. Absent = no fill. */
  shading?: string;
  /** Resolved per-edge borders; absent = renderer's default light grid. */
  borders?: CellBorders;
  /** Inner content box (absolute coords) the renderer clips cell content to, so
   *  over-wide text never paints onto the border or into the neighbour column —
   *  matches Word. Horizontal band only (full cell height) so descenders and
   *  rowspan-straddling content are never cut. */
  contentClip?: { x: number; y: number; width: number; height: number };
}

export interface PlacedTableRow {
  y: number;
  height: number;
  cells: PlacedTableCell[];
}

export interface PlacedTable {
  x: number;
  y: number;
  width: number;
  height: number;
  rows: PlacedTableRow[];
  /** Per-column widths — column boundary hit-testing reads these (merged cells
   *  make row-cell edges unreliable as column markers). */
  colWidths: number[];
}

export interface PlacedBlock {
  blockId: string;
  x: number; // content-box position on the page
  y: number;
  /** For paragraphs split across pages: which line range of the block lives here. */
  firstLineIndex: number;
  lines: LineBox[];
  /** List marker ("•", "3.", "b)") — paint-only, drawn in the hanging indent on
   *  the first line. Never measured: markers cannot change line breaking. */
  marker?: { text: string; style: CharStyle; x: number };
  /** TOC entry decoration: the target's page number right-aligned at numX with
   *  a dot leader, on line `lineIndex` of this chunk. Paint-only — resolved in
   *  an engine post-pass from the final page map, so it is never stale.
   *  `targetId` is the heading block this entry points at (PDF emits a GoTo link). */
  toc?: { numText: string; numX: number; lineIndex: number; style: CharStyle; targetId: string };
  /** Present when this placed block is an image / table / equation (lines stays empty). */
  image?: PlacedImage;
  table?: PlacedTable;
  equation?: PlacedEquation;
}

export interface Page {
  index: number;
  /** Displayed page number (honors section pageNumberStart) — what {page} shows
   *  in the footer and what "recalculate TOC" reads. Differs from `index` when a
   *  section restarts numbering. */
  number: number;
  blocks: PlacedBlock[];
  /** Per-page dimensions — sections can change page size/margins mid-document,
   *  so paint and hit-testing must read THESE, not the tree-level defaults. */
  widthPx: number;
  heightPx: number;
  marginPx: { top: number; right: number; bottom: number; left: number };
  /** The body's content box edges. Differ from the margins when a tall
   *  header/footer pushes the body (band-edit boundary + band hit regions). */
  contentTopPx: number;
  contentBottomPx: number;
  /** w:background page fill ("#rrggbb"); absent = white. */
  pageColorHex?: string;
  /** w:pgBorders page border box; absent = no border. */
  pageBorders?: PageBorders;
  /** X positions (page coords) of column separator lines; absent = none. */
  columnSeparatorsX?: number[];
  /** Footnote separator rule (present only on pages carrying notes). */
  footnoteRuleY?: number;
  /** Margin-band stories, already positioned in page coords. Read-only: the
   *  geometry index deliberately skips them (no caret/selection in bands yet). */
  header?: PlacedBlock[];
  footer?: PlacedBlock[];
  /** WHICH model container produced each band on this page (first/even
   *  variants override the default) — story-edit nav reads the source list. */
  headerSource?: BandContainer;
  footerSource?: BandContainer;
}

export interface LayoutTree {
  pages: Page[];
  /** Dimensions of the FINAL section (`doc.section`) — defaults/fallbacks only;
   *  per-page truth lives on each Page. */
  pageWidthPx: number;
  pageHeightPx: number;
  marginPx: { top: number; right: number; bottom: number; left: number };
}
