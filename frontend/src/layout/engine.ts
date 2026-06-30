// Layer 2: layout engine — the ONLY module tree that imports @chenglou/pretext.
//
// Per paragraph:  runs -> RichInlineItem[] -> prepareRichInline()   (cached by revision)
//                 then layoutNextRichInlineLineRange() loop -> LineBox[]
// Then block flow + pagination -> LayoutTree with page-absolute coordinates.

import {
  layoutNextRichInlineLineRange,
  materializeRichInlineLineRange,
  type RichInlineCursor,
} from "@chenglou/pretext/rich-inline";
import type { Block, CharStyle, Document, EquationBlock, ImageBlock, LineNumbering, Paragraph, ParaStyle, Run, SectionBreakType, SectionPatch, SectionProps, TableBlock, TableCell, TabStop } from "@cw/shared";
import { effectiveFractions, isHiddenParagraph } from "@cw/shared";
import { formatListNumber, markerText, type ListDefinition, type ListLevel } from "@cw/shared";
import type { InlineFragment, LayoutTree, LineBox, Page, PlacedBlock, PlacedImage } from "./layoutTree";
import { PrepareCache, prepareRunSegment, setActiveCjkFallback, applyCjkLocale, type PreparedSegment } from "./prepareCache";
import { charStyleToFont, fontMetrics, measureTextWidth } from "./metrics";
import { widthScale, PARA_BORDER_PAD_PX } from "../paint/paintStyle";
import { baseLevelFor, effectiveAlign, hasRtlChars, levelsFor, visualOrder } from "./bidi";
import { setActiveFontRegistry, type CustomFontRegistry } from "../fonts/customRegistry";
import { equationBox, EQUATION_DISPLAY_PX } from "./math/equationLayout";
import type { MathBox } from "./math/mathBox";
import { MATH_FONT_FAMILY } from "../fonts/clones";

/** Display equations are typeset with the bundled math font (STIX Two Math). */
const eqBox = (b: EquationBlock): MathBox => equationBox(b.equation, MATH_FONT_FAMILY, EQUATION_DISPLAY_PX);

/** Word's default tab interval when a `\t` runs past the last explicit stop
 *  (0.5 inch at 96 dpi). */
const DEFAULT_TAB_STOP_PX = 48;

/** Active default tab interval for the current layout pass — set from the
 *  document's w:defaultTabStop (settings.xml) at the top of layout(), mirroring
 *  the await-free module-level config the CJK fallback/font registry use. Falls
 *  back to DEFAULT_TAB_STOP_PX for documents that don't override it. */
let activeDefaultTabStopPx = DEFAULT_TAB_STOP_PX;

/** Paint-only paragraph decoration (w:shd fill + w:pBdr border box) for a placed
 *  paragraph chunk. `boxWidth` is the content width between the paragraph's
 *  indents at this placement site; `height` is the chunk's rendered height.
 *  Returns undefined for an undecorated paragraph (no field set). Shared by every
 *  paragraph-placement path (body flow, floats, table cells, bands, footnotes) so
 *  borders/shading render consistently everywhere a paragraph lands. */
function paraDecorFor(style: ParaStyle, boxWidth: number, height: number): NonNullable<PlacedBlock["paraDecor"]> | undefined {
  const { borders, shading } = style;
  if (!borders && !shading) return undefined;
  // Border-to-text padding: the painters expand the box outward by `pad` so a
  // wide/double rule sits OUTSIDE the glyphs (and the shading fill reaches it).
  // Only borders need the gap; a shading-only paragraph keeps its text-edge box
  // so its fill is byte-identical to before.
  const pad = borders ? PARA_BORDER_PAD_PX : 0;
  return { width: Math.max(0, boxWidth), height, pad, ...(shading ? { shading } : {}), ...(borders ? { borders } : {}) };
}

/** Vertical space a bordered paragraph's box claims ABOVE its text, reserved in
 *  the block gap so the box (and any shading fill) can't paint over the adjacent
 *  block: the top rule width (if any) plus the border-to-text padding — the box
 *  grows by `pad` above the text whenever the paragraph has borders, regardless
 *  of which edges carry a rule. 0 for an unbordered paragraph. Used by both the
 *  normal and the float-adjacent placement paths. */
function paraBorderReserveTop(style: ParaStyle): number {
  if (!style.borders) return 0;
  return (style.borders.top?.widthPx ?? 0) + PARA_BORDER_PAD_PX;
}

/** Mirror of paraBorderReserveTop for the space below the text (bottom rule + pad). */
function paraBorderReserveBottom(style: ParaStyle): number {
  if (!style.borders) return 0;
  return (style.borders.bottom?.widthPx ?? 0) + PARA_BORDER_PAD_PX;
}

const sumLineHeights = (lines: LineBox[]): number => lines.reduce((h, l) => h + l.height, 0);

export interface LayoutOptions {
  /** Band being story-edited: rendered RAW ({page} tokens literal, real block
   *  ids) so geometry offsets align with the model during editing. */
  rawBand?: "header" | "footer" | null;
}

export interface LayoutEngine {
  /** Full or incremental relayout. Pass dirty ids from ApplyResult; omit for full. */
  layout(doc: Document, dirtyBlockIds?: string[], options?: LayoutOptions): LayoutTree;
  evict(blockId: string): void;
  /** Drop ALL cached layout — call when swapping in a wholly different document.
   *  The importer re-mints block ids from i0 at revision 0, so without a reset the
   *  new document's paragraphs hit the previous document's cached lines (same
   *  id+revision+width) and the two documents render merged. */
  reset(): void;
}

/** Per-engine CJK config, re-asserted at the top of every layout pass so it's never
 *  read across an `await` (concurrency-safe without a lock — see prepareCache.ts). */
export interface LayoutEngineOptions {
  /** CJK fallback family: script-split CJK runs render with it. `""`/whitespace or
   *  omitted = no fallback (browser system fallback renders CJK on screen). */
  cjkFallback?: string | null;
  /** pretext analyzer locale (CJK line-break tuning) for this engine's layouts. */
  cjkLocale?: string;
}

/** @param fontRegistry custom-font registry to make active for this engine's layout
 *  passes (editor mount / export job). Omit to leave the active registry untouched
 *  (tests + recalc paths that use the process-default global).
 *  @param opts per-engine CJK fallback + locale (defaults: no fallback, neutral). */
export function createLayoutEngine(fontRegistry?: CustomFontRegistry, opts?: LayoutEngineOptions): LayoutEngine {
  const cjkFallback = opts?.cjkFallback ?? null;
  const cjkLocale = opts?.cjkLocale;
  const prepCache = new PrepareCache();
  // Second cache tier: LineBox[] per (block revision, width). A keystroke
  // re-breaks ONE paragraph; every other block's lines are reused as-is and
  // re-layout degenerates to the pagination walk (pure arithmetic).
  const linesCache = new Map<string, { revision: number; width: number; lines: LineBox[] }>();

  // Third cache tier: the full measureTable result per (table revision, width).
  // measureTable re-walks every cell (1191 cells / ~9ms on a 98-page report) on
  // EVERY layout pass, even for unrelated edits. Its output depends only on the
  // table + the column width: any edit to the table OR a cell paragraph bumps
  // table.revision (replaceParagraphAt/replaceCellBlocks/replaceTable all do), and
  // the result is position-independent (placement assigns page coords afterwards),
  // so an unchanged table at the same width reuses its measurement verbatim.
  const tableCache = new Map<string, { revision: number; width: number; measured: ReturnType<typeof measureTable> }>();

  // Ids visited during the current FULL layout (null during a partial rawBand
  // pass, which only touches one band and must not evict the rest). After a full
  // layout we prune both caches to this set, so deleted blocks can't leak.
  let touched: Set<string> | null = null;

  const getLines = (p: Paragraph, width: number): LineBox[] => {
    touched?.add(p.id);
    const hit = linesCache.get(p.id);
    if (hit && hit.revision === p.revision && hit.width === width) return hit.lines;
    const lines = paragraphLines(p, width, prepCache);
    linesCache.set(p.id, { revision: p.revision, width, lines });
    return lines;
  };

  // Memoized top-level table measurement (body flow only; nested/band tables stay
  // direct — their width is derived inside an ancestor solve and their revision may
  // not bump independently).
  const getMeasuredTable = (
    t: TableBlock,
    width: number,
    getLinesF: (p: Paragraph, w: number) => LineBox[],
    listCtx: CellListCtx,
  ): ReturnType<typeof measureTable> => {
    touched?.add(t.id);
    const hit = tableCache.get(t.id);
    if (hit && hit.revision === t.revision && hit.width === width) return hit.measured;
    const measured = measureTable(t, width, getLinesF, listCtx);
    tableCache.set(t.id, { revision: t.revision, width, measured });
    return measured;
  };

  // Fourth cache tier: per-page header/footer layout. On a 98-page report this is
  // ~196 layoutBand calls (header+footer × page) re-shaped from scratch EVERY pass
  // — ~20ms, and a body keystroke never touches a band. layoutBand's output is a
  // pure function of (band blocks, width, originX, pageNum, pageCount, rawMode) +
  // the active font registry, and the result is shiftPlaced into fresh objects
  // before use, so cache it. Band blocks are keyed by array identity: a body edit
  // keeps doc.section.header's reference (cache hit); a band edit mints a new array
  // (miss). Page-count / page-number changes vary the key (correct: {numpages} /
  // {page} must re-render).
  const bandArrayId = new WeakMap<object, number>();
  let bandArraySeq = 0;
  const idOfBandArray = (blocks: Block[]): number => {
    let id = bandArrayId.get(blocks);
    if (id === undefined) { id = ++bandArraySeq; bandArrayId.set(blocks, id); }
    return id;
  };
  // {date}/{time} resolve against `new Date()` at layout time (not in the cache
  // key), so a band carrying them must never be cached — otherwise a footer clock
  // would freeze. Detected once per band-array identity.
  const bandVolatile = new WeakMap<object, boolean>();
  const hasVolatileToken = (blocks: Block[]): boolean => {
    let v = bandVolatile.get(blocks);
    if (v === undefined) {
      v = blockTexts(blocks).some((t) => /\{(date|time)\}/.test(t));
      bandVolatile.set(blocks, v);
    }
    return v;
  };
  const bandLayoutCache = new Map<string, { placed: PlacedBlock[]; height: number }>();
  let touchedBands: Set<string> | null = null;
  const getBandLayout = (
    blocks: Block[],
    width: number,
    originX: number,
    pageNum: number,
    pageCount: number,
    bandCache: PrepareCache,
    rawMode: boolean,
  ): { placed: PlacedBlock[]; height: number } => {
    if (hasVolatileToken(blocks)) return layoutBand(blocks, width, originX, pageNum, pageCount, bandCache, rawMode);
    const key = `${idOfBandArray(blocks)}|${width}|${originX}|${pageNum}|${pageCount}|${rawMode ? 1 : 0}`;
    touchedBands?.add(key);
    const hit = bandLayoutCache.get(key);
    if (hit) return hit;
    const res = layoutBand(blocks, width, originX, pageNum, pageCount, bandCache, rawMode);
    bandLayoutCache.set(key, res);
    return res;
  };

  return {
    layout(doc: Document, _dirtyBlockIds?: string[], options?: LayoutOptions): LayoutTree {
      // Assert this engine's font registry for the whole (synchronous) layout pass
      // so cloneFamilyFor/metricsFor resolve against the right custom fonts. Safe
      // for concurrent editors/jobs: layout() never awaits, so nothing can swap the
      // active registry mid-pass.
      if (fontRegistry) setActiveFontRegistry(fontRegistry);
      // Same await-free re-assert for this engine's CJK config (see prepareCache.ts):
      // scriptSplitRuns reads the active fallback during this synchronous pass.
      setActiveCjkFallback(cjkFallback);
      applyCjkLocale(cjkLocale);
      // Honor the document's w:defaultTabStop (settings.xml) for this pass; a `\t`
      // past the last explicit stop advances by this interval (Word's behavior).
      activeDefaultTabStopPx = doc.defaultTabStopPx && doc.defaultTabStopPx > 0 ? doc.defaultTabStopPx : DEFAULT_TAB_STOP_PX;
      const rawBand = options?.rawBand ?? null;
      touched = rawBand === null ? new Set<string>() : null;
      touchedBands = rawBand === null ? new Set<string>() : null;
      try {
        const tree = layoutDocument(doc, getLines, (p) => {
          touched?.add(p.id);
          return prepCache.get(p);
        }, rawBand, getMeasuredTable, getBandLayout);
        if (touched) {
          for (const id of linesCache.keys()) if (!touched.has(id)) linesCache.delete(id);
          for (const id of tableCache.keys()) if (!touched.has(id)) tableCache.delete(id);
          prepCache.retainOnly(touched);
        }
        if (touchedBands) {
          for (const k of bandLayoutCache.keys()) if (!touchedBands.has(k)) bandLayoutCache.delete(k);
        }
        return tree;
      } finally {
        touched = null;
        touchedBands = null;
      }
    },
    evict(blockId: string): void {
      prepCache.evict(blockId);
      linesCache.delete(blockId);
      tableCache.delete(blockId);
    },
    reset(): void {
      prepCache.clear();
      linesCache.clear();
      tableCache.clear();
      bandLayoutCache.clear();
    },
  };
}

// ---------------------------------------------------------------------------
// Paragraph -> LineBox[]

function countSpaces(text: string): number {
  let n = 0;
  for (let i = 0; i < text.length; i++) if (text.charCodeAt(i) === 0x20) n++;
  return n;
}

const isCollapsibleWS = (c: string | undefined): boolean =>
  c === " " || c === "\t" || c === "\n" || c === "\r";

/** Map a pretext fragment's text back onto the model run text. Pretext collapses
 *  consecutive whitespace (white-space: normal), so the rendered fragment can be
 *  SHORTER than its model range — a naive indexOf fails and every offset after
 *  the collapse drifts (caret lands on the wrong character; the "backspace
 *  deletes 'd' instead of 'o'" bug). The tolerant walk skips extra model
 *  whitespace and emits an offset map only when the mapping is non-identity. */
function mapFragmentToRun(
  runText: string,
  from: number,
  fragText: string,
): { start: number; end: number; map?: number[] } {
  const exact = runText.indexOf(fragText, from);
  if (exact >= 0) return { start: exact, end: exact + fragText.length };

  const map: number[] = [];
  let i = from;
  let j = 0;
  let start = -1;
  while (j < fragText.length && i < runText.length) {
    if (runText[i] === fragText[j]) {
      if (start < 0) start = i;
      map.push(i - start);
      i++;
      j++;
      continue;
    }
    if (isCollapsibleWS(runText[i])) {
      i++; // model whitespace with no rendered counterpart (collapsed or eaten)
      continue;
    }
    start = -1; // hard mismatch
    break;
  }
  if (start >= 0 && j === fragText.length) {
    // Trailing eaten whitespace belongs to the line break / next gap, not to
    // this fragment — the range ends at the last rendered character.
    map.push(i - start);
    const identity = i - start === fragText.length;
    return identity ? { start, end: i } : { start, end: i, map };
  }
  // Last resort (should not happen): keep the old advancing fallback.
  return { start: from, end: from + fragText.length };
}

interface RawFrag {
  frag: InlineFragment;
  hadGap: boolean; // eaten separator whitespace preceded this fragment
  gap: number; // width (px) of that separator gap — needed to rebuild x after bidi reorder
  spaces: number; // U+0020 count inside the fragment text
}

interface BrokenLine {
  frags: RawFrag[]; // fragments with line-LOCAL x (no alignment applied)
  width: number;
  height: number;
  ascent: number;
  end: RichInlineCursor;
}

function segRunOffsets(runs: Run[], base: number): number[] {
  const runStarts: number[] = [];
  let acc = base;
  for (const run of runs) {
    runStarts.push(acc);
    acc += run.text.length;
  }
  return runStarts;
}

/** A fully-hidden paragraph takes ZERO space — unless it carries a structural
 *  break, which we keep so the break isn't lost. */
function isHiddenLayoutSkip(p: Paragraph): boolean {
  return isHiddenParagraph(p) && !p.style.sectionBreak && !p.style.pageBreakBefore;
}

/** Line-box height for ONE line, honoring fixed line spacing (docx w:lineRule).
 *  "exact" pins the box to `lineHeightPx` (taller glyphs clip); "atLeast" floors
 *  it there but lets a taller line grow; "auto" (the default) scales the dominant
 *  font size by the multiplier. `natural` is the line's glyph height (ascent +
 *  descent of its tallest run). */
function lineBoxHeight(style: ParaStyle, fontSizePx: number, natural: number): number {
  if (style.lineRule === "exact") return style.lineHeightPx ?? natural;
  if (style.lineRule === "atLeast") return Math.max(natural, style.lineHeightPx ?? 0);
  return Math.max(natural, style.lineHeight * fontSizePx);
}

/** Default line metrics for a fragment-less line (empty paragraph / segment). */
function emptyLineMetrics(p: Paragraph): { height: number; ascent: number } {
  const style = p.runs[0]?.style;
  const fontSize = style?.fontSizePx ?? 16;
  const m = style ? fontMetrics(charStyleToFont(style)) : { ascent: fontSize * 0.8, descent: fontSize * 0.2 };
  const natural = m.ascent + m.descent;
  const height = lineBoxHeight(p.style, fontSize, natural);
  return { height, ascent: (height - natural) / 2 + m.ascent };
}

/** Break ONE line of a SEGMENT at maxWidth. Mutates runCursors (snapshot before
 *  calling if the line might be rejected and re-broken at another width). */
function breakNextLine(
  p: Paragraph,
  seg: PreparedSegment,
  maxWidth: number,
  cursor: RichInlineCursor | undefined,
  runStarts: number[],
  runCursors: number[],
): BrokenLine | null {
  const range = layoutNextRichInlineLineRange(seg.prepared, maxWidth, cursor);
  if (range === null) return null;
  const line = materializeRichInlineLineRange(seg.prepared, range);

  const frags: RawFrag[] = [];
  let x = 0;
  let maxAscent = 0;
  let maxDescent = 0;
  let maxFontSize = 0;

  for (const f of line.fragments) {
    x += f.gapBefore;
    const run = seg.runs[f.itemIndex]!;
    // Inline equation: the math box drives the line's vertical metrics (not the
    // sentinel font), and the box rides on the fragment for paint/geometry.
    if (run.style.equation) {
      const box = equationBox(run.style.equation, MATH_FONT_FAMILY, run.style.fontSizePx);
      maxAscent = Math.max(maxAscent, box.ascent);
      maxDescent = Math.max(maxDescent, box.descent);
      maxFontSize = Math.max(maxFontSize, run.style.fontSizePx);
      const m = mapFragmentToRun(run.text, runCursors[f.itemIndex]!, f.text);
      runCursors[f.itemIndex] = m.end;
      frags.push({
        frag: {
          blockId: p.id,
          startOffset: runStarts[f.itemIndex]! + m.start,
          endOffset: runStarts[f.itemIndex]! + m.end,
          text: f.text,
          style: run.style,
          x,
          width: f.occupiedWidth,
          equation: box,
        },
        hadGap: f.gapBefore > 0,
        gap: f.gapBefore,
        spaces: 0,
      });
      x += f.occupiedWidth;
      continue;
    }
    const fm = fontMetrics(charStyleToFont(run.style));
    maxAscent = Math.max(maxAscent, fm.ascent);
    maxDescent = Math.max(maxDescent, fm.descent);
    maxFontSize = Math.max(maxFontSize, run.style.fontSizePx);
    // Character width scaling (w:w): pretext measures the natural advance; the run
    // occupies that × the scale, so reserve the scaled width here and advance by it.
    // scale === 1 (the common case) leaves occupiedWidth byte-identical.
    const occupied = f.occupiedWidth * widthScale(run.style);
    if (f.text.length > 0) {
      const m = mapFragmentToRun(run.text, runCursors[f.itemIndex]!, f.text);
      runCursors[f.itemIndex] = m.end;
      const frag: InlineFragment = {
        blockId: p.id,
        startOffset: runStarts[f.itemIndex]! + m.start,
        endOffset: runStarts[f.itemIndex]! + m.end,
        text: f.text,
        style: run.style,
        x,
        width: occupied,
      };
      if (m.map) frag.offsetMap = m.map;
      frags.push({ frag, hadGap: f.gapBefore > 0, gap: f.gapBefore, spaces: countSpaces(f.text) });
    }
    x += occupied;
  }

  if (maxFontSize === 0) {
    const style = p.runs[0]?.style;
    if (style) {
      const fm = fontMetrics(charStyleToFont(style));
      maxAscent = fm.ascent;
      maxDescent = fm.descent;
      maxFontSize = style.fontSizePx;
    }
  }

  const natural = maxAscent + maxDescent;
  const height = lineBoxHeight(p.style, maxFontSize, natural);
  return { frags, width: line.width, height, ascent: lineBaseline(p.style.textAlignment, height, maxAscent, maxDescent), end: line.end };
}

/** Baseline offset from the line-box top for the line's shared baseline, per the
 *  paragraph's w:textAlignment. "baseline"/absent and "center" both center the
 *  natural text block via half-leading (Word's default); "top"/"bottom" hug the
 *  respective edge of a tall line box. Mixed-size runs still share one baseline —
 *  this shifts that baseline within the box rather than aligning each run's edge. */
function lineBaseline(
  textAlignment: ParaStyle["textAlignment"],
  height: number,
  maxAscent: number,
  maxDescent: number,
): number {
  switch (textAlignment) {
    case "top":
      return maxAscent;
    case "bottom":
      return height - maxDescent;
    default:
      return (height - (maxAscent + maxDescent)) / 2 + maxAscent; // center / baseline
  }
}

interface RawLine {
  frags: RawFrag[];
  width: number;
  indent: number;
  lineMaxWidth: number;
  height: number;
  ascent: number;
  lastOfSegment: boolean;
  emptyOffset?: number;
  leaders?: LineBox["leaders"];
  tabArrows?: LineBox["tabArrows"];
  paragraphEnd?: boolean;
  lineBreak?: boolean;
  /** Tab-stop line with absolute fragment x baked in — never bidi-reordered. */
  tabbed?: boolean;
}

/** Next tab stop strictly past `curX`: the first explicit stop, else the next
 *  default-interval multiple. */
function resolveTabStop(curX: number, stops: TabStop[], contentWidth: number): TabStop {
  for (const s of stops) if (s.posPx > curX + 0.5) return s;
  const interval = activeDefaultTabStopPx;
  let pos = (Math.floor(curX / interval) + 1) * interval;
  if (pos <= curX) pos = curX + interval;
  return { posPx: Math.min(pos, Math.max(curX + 1, contentWidth)), align: "left", leader: "none" };
}

/** Width of a piece's text up to (and including) its decimal separator — for
 *  decimal-aligned tab stops. Measured with the piece's leading run font. */
function decimalPrefixWidth(runs: Run[]): number {
  const text = runs.map((r) => r.text).join("");
  const font = charStyleToFont(runs[0]?.style ?? DEFAULT_DECIMAL_STYLE);
  // The decimal separator is "." (en-US); fall back to "," only if there's no
  // period, so a thousands comma in "$1,234.50" doesn't hijack the alignment.
  let m = text.indexOf(".");
  if (m < 0) m = text.indexOf(",");
  return measureTextWidth(m < 0 ? text : text.slice(0, m + 1), font);
}
const DEFAULT_DECIMAL_STYLE = { fontFamily: "Georgia, serif", fontSizePx: 16 } as CharStyle;

/** Paint style for margin line numbers (w:lnNumType) — small, muted, like Word. */
const LINE_NUMBER_STYLE = { fontFamily: "Georgia, serif", fontSizePx: 12, color: "#9aa0a6" } as CharStyle;
/** Default gap (px) between the line numbers and the text edge when w:distance is absent. */
const LINE_NUMBER_DISTANCE = 18;

/** Lay out a `\t`-containing segment. Each tab-delimited piece is positioned at
 *  the next tab stop (left/center/right/decimal, with optional leaders); the
 *  piece's text then WRAPS normally — its first line continues from the stop and
 *  any overflow flows to new full-width lines at the left margin (so a common
 *  leading-tab first-line indent wraps like Word instead of running off the
 *  page). Returns one RawLine per visual line. */
function layoutTabbedSegment(
  p: Paragraph,
  segRuns: Run[],
  segStart: number,
  contentWidth: number,
  firstIndent: number,
): RawLine[] {
  // Split runs at "\t" into pieces, tracking each piece's global start offset
  // (the "\t" itself occupies one offset — caret can land between pieces).
  const pieces: { runs: Run[]; start: number }[] = [];
  let cur: Run[] = [];
  let pieceStart = segStart;
  let pos = segStart;
  for (const r of segRuns) {
    let local = 0;
    for (;;) {
      const idx = r.text.indexOf("\t", local);
      if (idx < 0) break;
      const piece = r.text.slice(local, idx);
      if (piece.length > 0) cur.push({ text: piece, style: r.style });
      pos += idx - local;
      pieces.push({ runs: cur, start: pieceStart });
      pos += 1;
      pieceStart = pos;
      cur = [];
      local = idx + 1;
    }
    const rest = r.text.slice(local);
    if (rest.length > 0) cur.push({ text: rest, style: r.style });
    pos += rest.length;
  }
  pieces.push({ runs: cur, start: pieceStart });

  const stops = (p.style.tabStops ?? []).slice().sort((a, b) => a.posPx - b.posPx);
  const baseStyle = segRuns.find((r) => r.text.length > 0)?.style ?? p.runs[0]?.style;
  const out: RawLine[] = [];

  // The current visual line being assembled.
  let curFrags: RawFrag[] = [];
  let curLeaders: NonNullable<RawLine["leaders"]> = [];
  let curArrows: NonNullable<RawLine["tabArrows"]> = [];
  let curX = firstIndent;
  let maxH = 0;
  let maxA = 0;
  let hasTab = false;

  const flush = (isLast: boolean): void => {
    if (maxH === 0 && baseStyle) {
      const fm = fontMetrics(charStyleToFont(baseStyle));
      maxH = lineBoxHeight(p.style, baseStyle.fontSizePx, fm.ascent + fm.descent);
      maxA = (maxH - (fm.ascent + fm.descent)) / 2 + fm.ascent;
    }
    out.push({
      frags: curFrags,
      width: curX,
      indent: 0, // fragment x is already absolute (tab/indent baked in)
      lineMaxWidth: contentWidth,
      height: maxH,
      ascent: maxA,
      // A tab-positioned line stays ragged (its manual x must not be re-justified);
      // pure wrapped continuation lines justify normally when the paragraph does.
      lastOfSegment: isLast || hasTab,
      ...(hasTab ? { tabbed: true } : {}),
      ...(curLeaders.length > 0 ? { leaders: curLeaders } : {}),
      ...(curArrows.length > 0 ? { tabArrows: curArrows } : {}),
    });
    curFrags = [];
    curLeaders = [];
    curArrows = [];
    curX = 0;
    maxH = 0;
    maxA = 0;
    hasTab = false;
  };

  pieces.forEach((pc, i) => {
    // Position this piece at its tab stop (relative to the running x cursor).
    if (i > 0) {
      const stop = resolveTabStop(curX, stops, contentWidth);
      const align = stop.align ?? "left";
      const fullW = measureTextWidth(pc.runs.map((r) => r.text).join(""), charStyleToFont(pc.runs[0]?.style ?? baseStyle ?? DEFAULT_DECIMAL_STYLE));
      const fits = curX + fullW <= contentWidth; // right/center/decimal need the piece on one line
      let targetX = stop.posPx;
      if (fits && align === "right") targetX = stop.posPx - fullW;
      else if (fits && align === "center") targetX = stop.posPx - fullW / 2;
      else if (fits && align === "decimal") targetX = stop.posPx - decimalPrefixWidth(pc.runs);
      if (targetX < curX) targetX = curX;
      if ((stop.leader ?? "none") !== "none" && targetX > curX + 1 && baseStyle) {
        curLeaders.push({ x1: curX, x2: targetX, kind: stop.leader ?? "none", color: baseStyle.color, fontSizePx: baseStyle.fontSizePx });
      }
      // Formatting-marks overlay: remember the gap this tab opened (drawn as "→").
      curArrows.push({ x1: curX, x2: Math.max(targetX, curX) });
      curX = targetX;
      hasTab = true;
    }
    if (pc.runs.length === 0) return; // empty piece (e.g. text before a leading tab)

    // Lay the piece out with wrapping: first sub-line continues from curX; any
    // overflow opens new full-width lines at the left margin. Use the CJK-split
    // runs the prepare actually saw so itemIndex/offset/cursor stay in sync.
    const prep = prepareRunSegment(pc.runs);
    const pseg: PreparedSegment = { prepared: prep.prepared, runs: prep.runs, startOffset: pc.start };
    const runStarts = segRunOffsets(prep.runs, pc.start);
    const runCursors = prep.runs.map(() => 0);
    let cursor: RichInlineCursor | undefined = undefined;
    let firstSub = true;
    for (;;) {
      const maxW = Math.max(24, contentWidth - (firstSub ? curX : 0));
      const bl = breakNextLine(p, pseg, maxW, cursor, runStarts, runCursors);
      if (bl === null) break;
      if (!firstSub) {
        flush(false); // previous line complete; this piece wrapped to a new line
      }
      const shift = firstSub ? curX : 0;
      for (const rf of bl.frags) {
        rf.frag.x += shift;
        curFrags.push(rf);
      }
      curX = shift + bl.width;
      maxH = Math.max(maxH, bl.height);
      maxA = Math.max(maxA, bl.ascent);
      cursor = bl.end;
      firstSub = false;
    }
  });

  flush(true);
  return out;
}

/** Justify a line: distribute `slack` px across every inter-word space (painted
 *  via ctx.wordSpacing, so the fragment stays ONE fillText) and every eaten
 *  inter-fragment gap. Mutates each fragment's x / width / wordSpacing in place;
 *  the caller applies any left-edge offset (indent / float box) separately. */
function justifyLine(frags: RawFrag[], slack: number): void {
  if (slack <= 0) return;
  const totalGaps = frags.reduce((s, rf) => s + rf.spaces + (rf.hadGap ? 1 : 0), 0);
  if (totalGaps <= 0) return;
  const extra = slack / totalGaps;
  let shift = 0;
  for (const rf of frags) {
    if (rf.hadGap) shift += extra;
    rf.frag.x += shift;
    if (rf.spaces > 0) {
      rf.frag.wordSpacingPx = extra;
      rf.frag.width += rf.spaces * extra;
      shift += rf.spaces * extra;
    }
  }
}

/** Per-code-unit bidi embedding levels over the WHOLE paragraph text (global
 *  offsets — fragments carry the same), or null when the paragraph is purely its
 *  base direction and needs no reordering (the LTR fast path AND a pure-Latin RTL
 *  paragraph, which only needs right-alignment, handled by effectiveAlign). */
function paragraphLevels(p: Paragraph): Int8Array | null {
  const base = baseLevelFor(p.style.direction);
  let text = "";
  const forced: [number, number][] = [];
  for (const r of p.runs) {
    const start = text.length;
    text += r.text;
    if (r.style.rtl) forced.push([start, text.length]);
  }
  if (base === 0 && forced.length === 0 && !hasRtlChars(text)) return null;
  let levels = levelsFor(text, base);
  if (levels === null) {
    // No strong-RTL characters resolved (e.g. an RTL paragraph of Latin text):
    // alignment alone makes it right-aligned, so only an explicit run `rtl` flag
    // forces any reordering.
    if (forced.length === 0) return null;
    levels = new Int8Array(text.length); // base-LTR fill; forced ranges go RTL below
  }
  for (const [a, b] of forced) {
    for (let i = a; i < b; i++) if ((levels[i]! & 1) === 0) levels[i] = levels[i]! + 1;
  }
  return levels;
}

/** Split a fragment whose model range spans more than one embedding level into
 *  per-level sub-fragments (each becomes a single-direction fillText). Skips the
 *  collapsed-whitespace case (offsetMap present), tagging the whole fragment with
 *  its start level — mixed-script inside a single collapsed run is vanishingly
 *  rare. Returns sub-RawFrags in LOGICAL order; only the first keeps the gap. */
function splitFragByLevel(rf: RawFrag, levels: Int8Array): RawFrag[] {
  const { frag } = rf;
  const s = frag.startOffset;
  const e = frag.endOffset;
  const startLevel = levels[s] ?? 0;
  // Keep collapsed-whitespace fragments and field results (PAGE/{page} tokens are
  // resolved as ONE fragment by a later post-pass — splitting "{page}" across the
  // brace/letter level boundary would stop the token from ever matching) atomic.
  if (frag.offsetMap || frag.style.fieldId !== undefined || frag.text.length !== e - s) {
    frag.level = startLevel;
    return [rf];
  }
  // Find level-change boundaries within [s, e).
  const bounds: number[] = [0];
  for (let i = s + 1; i < e; i++) if (levels[i] !== levels[i - 1]) bounds.push(i - s);
  if (bounds.length === 1) {
    frag.level = startLevel;
    return [rf];
  }
  bounds.push(e - s);
  const font = charStyleToFont(frag.style);
  const fullW = measureTextWidth(frag.text, font) || frag.width;
  const out: RawFrag[] = [];
  for (let k = 0; k < bounds.length - 1; k++) {
    const a = bounds[k]!;
    const b = bounds[k + 1]!;
    const sub = frag.text.slice(a, b);
    // Proportional width keeps the sub-pieces summing to the fragment's measured
    // width (pretext's occupiedWidth includes letter-spacing this measure omits).
    // Guard fullW===0 (zero-width control/combining runs) so we never emit NaN.
    const w = fullW > 0 ? (measureTextWidth(sub, font) / fullW) * frag.width : 0;
    out.push({
      hadGap: k === 0 ? rf.hadGap : false,
      gap: k === 0 ? rf.gap : 0,
      spaces: countSpaces(sub),
      frag: {
        ...frag,
        text: sub,
        startOffset: s + a,
        endOffset: s + b,
        width: w,
        level: levels[s + a] ?? startLevel,
      },
    });
  }
  return out;
}

/** Reorder a line's fragments into VISUAL (left-to-right) order per UAX#9 L2 and
 *  rebuild their line-local x. Returns the reordered frags and the new line width.
 *  Pure-LTR lines never reach here (paragraphLevels returns null). */
function bidiReorderLine(frags: RawFrag[], levels: Int8Array): { frags: RawFrag[]; width: number } {
  const split: RawFrag[] = [];
  for (const rf of frags) split.push(...splitFragByLevel(rf, levels));
  const order = visualOrder(split.map((rf) => rf.frag.level ?? 0));
  const visual = order.map((i) => split[i]!);
  let x = 0;
  for (const rf of visual) {
    x += rf.gap;
    rf.frag.x = x;
    x += rf.frag.width;
  }
  return { frags: visual, width: x };
}

/** Mirror a tab-stopped line for an RTL base-direction paragraph. `layoutTabbedSegment`
 *  always lays tab stops out in LOGICAL-advance space (the cursor starts at the start
 *  edge and grows in reading order, stop positions measured from that edge). For an
 *  RTL paragraph the start edge is the RIGHT margin, so the line is reflected about the
 *  content box: a piece at logical distance `d` from the start edge ends up `d` from the
 *  right edge, and the tab cells fill right-to-left.
 *
 *  Reflecting fragments in their LOGICAL order also yields the correct VISUAL order for
 *  free (the first logical cell lands on the right), so no extra L2 reorder is needed.
 *  Embedding levels are assigned (from `levels`, splitting mixed-level fragments; else
 *  the RTL base) so the renderer anchors each fragment on the correct edge and caret /
 *  hit-testing take the bidi (non-monotonic-x) path. Leaders and tab arrows mirror too. */
function mirrorTabbedLineRtl(rl: RawLine, levels: Int8Array | null, contentWidth: number): void {
  const split: RawFrag[] = [];
  for (const rf of rl.frags) {
    if (levels) {
      // splitFragByLevel copies the parent x onto every sub-piece, so re-spread the
      // sub-pieces sequentially across the parent's box before the reflection below.
      let x = rf.frag.x;
      for (const sub of splitFragByLevel(rf, levels)) {
        sub.frag.x = x;
        x += sub.frag.width;
        split.push(sub);
      }
    } else {
      rf.frag.level = 1; // pure base-direction RTL line — every fragment is RTL
      split.push(rf);
    }
  }
  for (const rf of split) rf.frag.x = contentWidth - rf.frag.x - rf.frag.width;
  rl.frags = split;
  // The reflected fragments already span the content box anchored at the right edge,
  // so neutralize phase-2 alignment (slack === 0 → no further shift).
  rl.width = contentWidth;
  if (rl.leaders) rl.leaders = rl.leaders.map((l) => ({ ...l, x1: contentWidth - l.x2, x2: contentWidth - l.x1 }));
  if (rl.tabArrows) rl.tabArrows = rl.tabArrows.map((a) => ({ x1: contentWidth - a.x2, x2: contentWidth - a.x1 }));
}

function paragraphLines(p: Paragraph, contentWidth: number, cache: PrepareCache): LineBox[] {
  const segments = cache.get(p);
  const lines: LineBox[] = [];
  const levels = paragraphLevels(p);
  const dir = p.style.direction;

  // Phase 1: break lines per SEGMENT (segments = soft-break "\v" pieces).
  // Fragments are line-LOCAL; justification needs to know each segment's last
  // line (it stays ragged, like a paragraph's final line), so alignment is a
  // second pass.
  const raw: RawLine[] = [];

  let first = true; // first line of the PARAGRAPH (first-line indent)
  // Index of the last raw line each segment contributed — the segment boundary
  // where a formatting mark lands (¶ for the paragraph's end, ↵ for a soft break
  // between segments). Set after every segment so each one always has an entry.
  const segLastLine: number[] = [];
  for (let si = 0; si < segments.length; si++) {
    const seg = segments[si]!;
    // Tab path: a segment with hard tabs lays out as one tab-stopped line (the
    // pretext collapse path would otherwise eat the tabs). Dormant unless run
    // text actually carries "\t".
    if (seg.runs.some((r) => r.text.includes("\t"))) {
      raw.push(...layoutTabbedSegment(p, seg.runs, seg.startOffset, contentWidth, first ? p.style.indentFirstLinePx : 0));
      first = false;
      segLastLine[si] = raw.length - 1;
      continue;
    }

    const runStarts = segRunOffsets(seg.runs, seg.startOffset);
    const runCursors: number[] = seg.runs.map(() => 0);
    let cursor: RichInlineCursor | undefined = undefined;
    let produced = false;

    for (;;) {
      const indent = first ? p.style.indentFirstLinePx : 0;
      const lineMaxWidth = contentWidth - indent;
      const bl = breakNextLine(p, seg, lineMaxWidth, cursor, runStarts, runCursors);
      if (bl === null) break;
      raw.push({
        frags: bl.frags,
        width: bl.width,
        indent,
        lineMaxWidth,
        height: bl.height,
        ascent: bl.ascent,
        lastOfSegment: false,
      });
      cursor = bl.end;
      first = false;
      produced = true;
    }

    if (!produced) {
      // Empty segment ("\v" at paragraph end, or "\v\v"): a blank line the
      // caret can land on — its offset is the segment start.
      const m = emptyLineMetrics(p);
      raw.push({
        frags: [],
        width: 0,
        indent: 0,
        lineMaxWidth: contentWidth,
        height: m.height,
        ascent: m.ascent,
        lastOfSegment: true,
        emptyOffset: seg.startOffset,
      });
      first = false;
    } else {
      raw[raw.length - 1]!.lastOfSegment = true;
    }
    segLastLine[si] = raw.length - 1;
  }

  // Formatting marks: each segment boundary carries a pilcrow (the paragraph's
  // final segment) or a line-break arrow (a soft return separating segments).
  for (let si = 0; si < segLastLine.length; si++) {
    const idx = segLastLine[si];
    if (idx === undefined) continue;
    if (si === segments.length - 1) raw[idx]!.paragraphEnd = true;
    else raw[idx]!.lineBreak = true;
  }

  // Phase 2: bidi reorder (RTL/mixed only), alignment + justification, final boxes.
  const align = effectiveAlign(p.style.align, dir);
  const rtlBase = baseLevelFor(dir) === 1;
  let y = 0;
  for (const rl of raw) {
    // Tabbed lines bake absolute x at tab stops, so the generic edge-to-edge bidi
    // reorder (which discards those x positions) can't run on them. An RTL paragraph
    // instead REFLECTS the line about the content box: tab stops then measure from
    // the right (start) edge and cells fill right-to-left, the Word w:bidi behavior.
    if (rl.tabbed) {
      // Mirror any RTL tabbed line that carries geometry — text, leaders, or just a
      // tab-arrow overlay (a "\t"-only line has arrows but no fragments/leaders).
      const hasGeometry =
        rl.frags.length > 0 || (rl.leaders?.length ?? 0) > 0 || (rl.tabArrows?.length ?? 0) > 0;
      if (rtlBase && hasGeometry) mirrorTabbedLineRtl(rl, levels, contentWidth);
    } else if (levels !== null && rl.frags.length > 0) {
      const r = bidiReorderLine(rl.frags, levels);
      rl.frags = r.frags;
      rl.width = r.width;
    }
    const slack = rl.lineMaxWidth - rl.width;
    let startX = rl.indent;
    if (align === "center") startX += slack / 2;
    else if (align === "right") startX += slack;
    // Justify works on the VISUAL fragment order (bidi reorder already ran), so it
    // distributes slack across the line's gaps edge-to-edge regardless of direction.
    else if (align === "justify" && !rl.lastOfSegment) justifyLine(rl.frags, slack);
    for (const rf of rl.frags) rf.frag.x += startX;
    const box: LineBox = { y, height: rl.height, ascent: rl.ascent, fragments: rl.frags.map((rf) => rf.frag) };
    if (rl.emptyOffset !== undefined) box.emptyOffset = rl.emptyOffset;
    if (rl.leaders && rl.leaders.length > 0) {
      box.leaders = rl.leaders.map((l) => ({ ...l, x1: l.x1 + startX, x2: l.x2 + startX }));
    }
    if (rl.tabArrows && rl.tabArrows.length > 0) {
      box.tabArrows = rl.tabArrows.map((a) => ({ x1: a.x1 + startX, x2: a.x2 + startX }));
    }
    if (rl.paragraphEnd) box.paragraphEnd = true;
    if (rl.lineBreak) box.lineBreak = true;
    lines.push(box);
    y += rl.height;
  }

  return lines;
}

// ---------------------------------------------------------------------------
// Block flow + pagination (line-level page breaks: a paragraph may span pages).
// Word-default break rules: orphan control (≥2 lines stay at the bottom of the
// page where a paragraph starts, else the whole paragraph moves), widow control
// (≥2 lines reach the paragraph's final page), keep-with-next for headings.

const ORPHAN_MIN = 2; // min lines kept on the starting page when breaking
const WIDOW_MIN = 2; // min lines carried onto the paragraph's last page
const TOC_GUTTER = 48; // right gutter reserved on TOC entries for the page number

// ---------------------------------------------------------------------------
// Sections: a section-break paragraph TERMINATES a section (OOXML sectPr
// placement); blocks after the last break belong to `doc.section`. Absent
// patch fields inherit from doc.section ("link to previous").

const BAND_KEYS = ["header", "footer", "headerFirst", "headerEven", "footerFirst", "footerEven"] as const;

export function effectiveSection(base: SectionProps, patch: SectionPatch): SectionProps {
  const out: SectionProps = {
    pageWidthPx: patch.pageWidthPx ?? base.pageWidthPx,
    pageHeightPx: patch.pageHeightPx ?? base.pageHeightPx,
    marginPx: patch.marginPx ?? base.marginPx,
  };
  // columns: undefined = inherit, null = explicitly single-column
  const columns = patch.columns === undefined ? base.columns : (patch.columns ?? undefined);
  if (columns) out.columns = columns;
  // page-number restart is a section's OWN property — never inherited
  if (patch.pageNumberStart !== undefined) out.pageNumberStart = patch.pageNumberStart;
  // Band distances inherit from the document section (the report sets them once).
  const headerDist = patch.headerDistancePx ?? base.headerDistancePx;
  if (headerDist !== undefined) out.headerDistancePx = headerDist;
  const footerDist = patch.footerDistancePx ?? base.footerDistancePx;
  if (footerDist !== undefined) out.footerDistancePx = footerDist;
  // Page fill & borders inherit from the document section unless overridden.
  const pageColorHex = patch.pageColorHex ?? base.pageColorHex;
  if (pageColorHex !== undefined) out.pageColorHex = pageColorHex;
  const pageBorders = patch.pageBorders ?? base.pageBorders;
  if (pageBorders !== undefined) out.pageBorders = pageBorders;
  // Line numbering is a section's OWN property (like page-number restart): a
  // section either declares its own w:lnNumType or has none — it never inherits.
  if (patch.lineNumbering !== undefined) out.lineNumbering = patch.lineNumbering;
  for (const key of BAND_KEYS) {
    const blocks = patch[key] ?? base[key];
    if (blocks) out[key] = blocks;
  }
  return out;
}

/** Top Y of the header band. Word anchors the header's TOP edge at w:header from
 *  the page top (the band grows down); absent that, center it in the top margin. */
function headerBandTop(s: SectionProps, bandHeight: number): number {
  if (s.headerDistancePx !== undefined) return s.headerDistancePx;
  return Math.max(8, (s.marginPx.top - bandHeight) / 2);
}

/** Top Y of the footer band. Word anchors the footer's BOTTOM edge at w:footer
 *  from the page bottom (the band grows up), so a tall footer sits higher than a
 *  bottom-flush one; absent that, center it in the bottom margin. */
function footerBandTop(s: SectionProps, bandHeight: number): number {
  if (s.footerDistancePx !== undefined) return s.pageHeightPx - s.footerDistancePx - bandHeight;
  return s.pageHeightPx - Math.max(8, (s.marginPx.bottom - bandHeight) / 2) - bandHeight;
}

interface ResolvedSection {
  props: SectionProps;
  /** Index (inclusive) of this section's last top-level block. */
  endBlock: number;
  /** The OOXML w:type that governs how THIS section's first page begins
   *  ("nextPage"/"evenPage"/"oddPage"). Read when the engine starts the section. */
  breakType: SectionBreakType;
}

export function resolveSections(doc: Document): ResolvedSection[] {
  const out: ResolvedSection[] = [];
  for (let i = 0; i < doc.blocks.length; i++) {
    const b = doc.blocks[i]!;
    if (b.kind === "paragraph" && b.style.sectionBreak) {
      out.push({
        props: effectiveSection(doc.section, b.style.sectionBreak.props),
        endBlock: i,
        breakType: b.style.sectionBreak.type,
      });
    }
  }
  // The trailing body section keeps its own start type (even/odd parity), so a
  // document whose final section begins on a parity page is honored, not flattened.
  out.push({ props: doc.section, endBlock: doc.blocks.length - 1, breakType: doc.section.breakType ?? "nextPage" });
  return out;
}

function layoutDocument(
  doc: Document,
  getLines: (p: Paragraph, width: number) => LineBox[],
  getPrepared: (p: Paragraph) => PreparedSegment[],
  rawBand: "header" | "footer" | null,
  measureTableCached: (
    t: TableBlock,
    width: number,
    getLinesF: (p: Paragraph, w: number) => LineBox[],
    listCtx: CellListCtx,
  ) => ReturnType<typeof measureTable> = measureTable,
  layoutBandCached: (
    blocks: Block[],
    width: number,
    originX: number,
    pageNum: number,
    pageCount: number,
    bandCache: PrepareCache,
    rawMode: boolean,
  ) => { placed: PlacedBlock[]; height: number } = layoutBand,
): LayoutTree {
  // Tall bands PUSH the content box (Word): a header taller than its margin
  // moves contentTop down; a tall footer raises contentBottom. Heights are
  // estimated per section by pre-laying each band variant once with
  // representative page numbers (the real per-page band layout happens after
  // the walk — actual heights can only be ≤ the estimate's digit width).
  const BAND_GAP = 6;
  const bandBoxCache = new Map<SectionProps, { top: number; bottom: number }>();
  let bandProbeCache: PrepareCache | null = null;
  const contentBoxOf = (s: SectionProps): { top: number; bottom: number } => {
    const hit = bandBoxCache.get(s);
    if (hit) return hit;
    let top = s.marginPx.top;
    let bottom = s.pageHeightPx - s.marginPx.bottom;
    const headers = [s.header, s.headerFirst, s.headerEven].filter((b): b is Block[] => !!b);
    const footers = [s.footer, s.footerFirst, s.footerEven].filter((b): b is Block[] => !!b);
    if (headers.length > 0 || footers.length > 0) {
      bandProbeCache ??= new PrepareCache();
      const cw = s.pageWidthPx - s.marginPx.left - s.marginPx.right;
      const cx = s.marginPx.left;
      const heightOf = (blocks: Block[]): number =>
        layoutBand(blocks, cw, cx, 999, 999, bandProbeCache!, false).height;
      // Height from the band's FIRST visible row to its bottom — leading blank
      // lines don't need exclusive body space (they overlap the body's blank tail).
      const bandVisibleHeight = (blocks: Block[]): number => {
        const band = layoutBand(blocks, cw, cx, 999, 999, bandProbeCache!, false);
        for (const pb of band.placed) {
          const visible =
            !!pb.table ||
            !!pb.image ||
            !!pb.equation ||
            pb.lines.some((l) => l.fragments.some((f) => f.text.trim().length > 0 || !!f.equation));
          if (visible) return band.height - pb.y;
        }
        return 0;
      };
      const headerH = headers.reduce((m, b) => Math.max(m, heightOf(b)), 0);
      if (headerH > 0) {
        const y0 = headerBandTop(s, headerH);
        top = Math.max(top, y0 + headerH + BAND_GAP);
      }
      const footerH = footers.reduce((m, b) => Math.max(m, heightOf(b)), 0);
      if (footerH > 0) {
        if (s.footerDistancePx !== undefined) {
          // Anchored footer: reserve down to its first VISIBLE row, not the band
          // top. Leading blank footer lines may overlap the body's (also blank)
          // tail — reserving them would strand a trailing blank line on a new page.
          const footerBottom = s.pageHeightPx - s.footerDistancePx;
          const visibleH = footers.reduce((m, b) => Math.max(m, bandVisibleHeight(b)), 0);
          bottom = Math.min(bottom, footerBottom - visibleH);
        } else {
          bottom = Math.min(bottom, footerBandTop(s, footerH) - BAND_GAP);
        }
      }
      bottom = Math.max(bottom, top + 24); // degenerate giant bands: keep a sliver
    }
    const box = { top, bottom };
    bandBoxCache.set(s, box);
    return box;
  };

  // Page geometry is PER SECTION — mutable, swapped as the walk crosses
  // section boundaries. (Single-section documents never reassign.)
  const sections = resolveSections(doc);
  let sec: SectionProps = sections[0]!.props;
  let contentX = 0;
  let contentTop = 0;
  let contentBottom = 0;
  let contentWidth = 0;
  // Newspaper columns: flow fills column boxes left-to-right before paging.
  // colBoxes holds each column's x + width; colWidth is the single MEASURE width
  // (the narrowest box for unequal columns, so wrapped content never overflows —
  // a documented single-pass limitation). colSeparatorsX are the gap midpoints
  // when a separator line is requested.
  let colCount = 1;
  let colWidth = 0;
  let colBoxes: { x: number; width: number }[] = [];
  let colSeparatorsX: number[] = [];
  const applySection = (props: SectionProps): void => {
    sec = props;
    const box = contentBoxOf(sec);
    contentX = sec.marginPx.left;
    contentTop = box.top;
    contentBottom = box.bottom;
    contentWidth = sec.pageWidthPx - sec.marginPx.left - sec.marginPx.right;
    colCount = Math.max(1, sec.columns?.count ?? 1);
    const gap = sec.columns?.gapPx ?? 24;
    const cols = sec.columns?.cols;
    if (colCount > 1 && cols && cols.length === colCount) {
      // Explicit per-column widths + per-column trailing gap (Word honors literally).
      let x = contentX;
      colBoxes = cols.map((c, i) => {
        const b = { x, width: c.widthPx };
        x += c.widthPx + (i < colCount - 1 ? c.spaceAfterPx : 0);
        return b;
      });
      colWidth = Math.min(...cols.map((c) => c.widthPx));
    } else if (colCount > 1) {
      colWidth = (contentWidth - (colCount - 1) * gap) / colCount;
      colBoxes = Array.from({ length: colCount }, (_, i) => ({
        x: contentX + i * (colWidth + gap),
        width: colWidth,
      }));
    } else {
      colWidth = contentWidth;
      colBoxes = [{ x: contentX, width: contentWidth }];
    }
    colSeparatorsX = sec.columns?.sep
      ? colBoxes.slice(0, -1).map((b, i) => (b.x + b.width + colBoxes[i + 1]!.x) / 2)
      : [];
  };
  applySection(sec);

  // List support: per-paragraph effective indent (style + level) and the
  // NUMBERING PASS — counters per (listId, level) in document order, emitting
  // paint-only markers. Counter state is pure string arithmetic; line caches
  // are untouched (markers can't change line breaking).
  const lists = doc.lists ?? {};
  const listLevelOf = (p: Paragraph): ListLevel | null => {
    const ref = p.style.list;
    if (!ref) return null;
    const def: ListDefinition | undefined = lists[ref.listId];
    if (!def) return null;
    return def.levels[Math.min(ref.level, def.levels.length - 1)] ?? null;
  };
  // Indents are mirrored for RTL paragraphs: the "start" indent (left + list) moves
  // to the RIGHT edge — which also opens the right-side gutter the list marker hangs
  // in — and the "end" indent (right) moves to the left. LTR is unchanged.
  const startIndentOf = (p: Paragraph): number => p.style.indentLeftPx + (listLevelOf(p)?.indentLeftPx ?? 0);
  const indentOf = (p: Paragraph): number =>
    p.style.direction === "rtl" ? (p.style.indentRightPx ?? 0) : startIndentOf(p);
  const rightIndentOf = (p: Paragraph): number =>
    p.style.direction === "rtl" ? startIndentOf(p) : (p.style.indentRightPx ?? 0);
  // The list-level indent alone (table cells add it ON TOP of the paragraph's
  // own indent, which their placement already applies).
  const listIndentOf = (p: Paragraph): number => listLevelOf(p)?.indentLeftPx ?? 0;

  const counters = new Map<string, number[]>();
  // Marker x is resolved at PLACEMENT time (placed.x - hangingPx) — with
  // columns, the numbering pass can't know which column the paragraph lands in.
  const markers = new Map<string, { text: string; style: CharStyle; hangingPx: number }>();
  const numberParagraph = (p: Paragraph): void => {
    const ref = p.style.list;
    const lvl = listLevelOf(p);
    if (!ref || !lvl) return;
    const def = lists[ref.listId]!;
    const stack = counters.get(ref.listId) ?? [];
    stack[ref.level] = (stack[ref.level] ?? lvl.start - 1) + 1;
    stack.length = ref.level + 1; // incrementing a level resets all deeper ones
    counters.set(ref.listId, stack);
    const text = markerText(def, ref.level, stack);
    if (text.length === 0) return;
    const baseStyle = p.runs[0]?.style;
    if (!baseStyle) return;
    markers.set(p.id, {
      text,
      style: { ...baseStyle, ...(lvl.markerStyle ?? {}) },
      hangingPx: lvl.hangingPx,
    });
  };

  // Numbering runs over the whole body in reading order BEFORE measurement —
  // recursing into table cells so a list continues straight through them (Word
  // counts 1,2,3 down a column of cells, not restarting per cell). Markers are
  // then read at placement time for both body and cell paragraphs.
  const numberBlocks = (blocks: Block[]): void => {
    for (const b of blocks) {
      if (b.kind === "paragraph") numberParagraph(b);
      else if (b.kind === "table") for (const row of b.rows) for (const cell of row.cells) numberBlocks(cell.blocks);
    }
  };
  numberBlocks(doc.blocks);
  const cellListCtx: CellListCtx = { indentOf: listIndentOf, markers };

  // Measure every block first (prepare/line caches make re-runs cheap); break
  // rules need full sizes before placement decisions.
  type Measured =
    | { kind: "para"; block: Paragraph; lines: LineBox[] }
    | { kind: "image"; block: ImageBlock; height: number }
    | { kind: "equation"; block: EquationBlock; box: MathBox; height: number }
    | { kind: "table"; block: TableBlock; rows: MeasuredRow[]; colWidths: number[]; height: number; tableWidth: number; xOffset: number };

  // Measured at each block's OWN section width (the section pointer advances
  // through the same block order the walk uses).
  const measured: Measured[] = [];
  let secIdx = 0;
  for (let bi = 0; bi < doc.blocks.length; bi++) {
    while (bi > sections[secIdx]!.endBlock) applySection(sections[++secIdx]!.props);
    const block = doc.blocks[bi]!;
    if (block.kind === "paragraph") {
      if (isHiddenLayoutSkip(block)) continue; // fully-hidden anchor: contributes no layout
      // Indented paragraphs (quotes, list levels) measure at the narrowed width;
      // colWidth IS contentWidth in single-column sections. TOC entries reserve
      // a right gutter so their text never collides with the page number.
      const gutter = block.style.tocEntry ? TOC_GUTTER : 0;
      measured.push({
        kind: "para",
        block,
        lines: getLines(block, colWidth - indentOf(block) - rightIndentOf(block) - gutter),
      });
    } else if (block.kind === "image") {
      // Anchored (out-of-flow) images contribute no flow height.
      measured.push({ kind: "image", block, height: block.anchor ? 0 : block.heightPx });
    } else if (block.kind === "equation") {
      const box = eqBox(block);
      measured.push({ kind: "equation", block, box, height: box.ascent + box.descent });
    } else {
      const t = measureTableCached(block, colWidth, getLines, cellListCtx);
      measured.push({ kind: "table", block, ...t });
    }
  }
  secIdx = 0;
  applySection(sections[0]!.props);

  // Which section governs each page — bands and paint read per-page geometry.
  const pageSections: SectionProps[] = [];
  const mkPage = (): Page => {
    pageSections.push(sec);
    const p: Page = {
      index: pageSections.length - 1,
      number: pageSections.length, // provisional; the post-pass sets the real value
      blocks: [],
      widthPx: sec.pageWidthPx,
      heightPx: sec.pageHeightPx,
      marginPx: sec.marginPx,
      contentTopPx: contentTop,
      contentBottomPx: contentBottom,
    };
    if (sec.pageColorHex !== undefined) p.pageColorHex = sec.pageColorHex;
    if (sec.pageBorders !== undefined) p.pageBorders = sec.pageBorders;
    if (colSeparatorsX.length) p.columnSeparatorsX = colSeparatorsX.slice();
    return p;
  };

  const pages: Page[] = [mkPage()];
  let page = pages[0]!;
  let y = contentTop;

  // TOC entries: last placed chunk per entry block — the page-number
  // decoration is attached there in a post-pass once page numbers exist.
  const tocPlacements = new Map<string, { chunk: PlacedBlock; rightEdge: number; para: Paragraph }>();

  // ---- footnotes ----------------------------------------------------------
  // A note reserves space at the bottom of the page its REFERENCE lands on:
  // `reserved` shrinks the effective bottom for everything placed after, and
  // the chunk carrying the ref is shrunk until chunk + notes co-fit (the
  // greedy version of Word's per-page fixpoint — one pass, deterministic).
  const FN_SEP = 14; // separator rule + padding above the first note
  const FN_INDENT = 18; // note text indent; the number marker hangs in it
  const fns = doc.footnotes ?? {};
  const hasNotes = Object.keys(fns).length > 0;
  let reserved = 0;
  const bottomY = (): number => contentBottom - reserved;
  const placedNotes = new Set<string>();
  const pageNotes: string[][] = [[]];
  const pageReserved: number[] = [0];
  const noteNumbers = new Map<string, string>(); // id -> marker text (the ref run's text)
  interface NoteMeasure {
    paras: { p: Paragraph; lines: LineBox[] }[];
    height: number;
  }
  const noteMeasure = (id: string): NoteMeasure | null => {
    const paras = fns[id];
    if (!paras || paras.length === 0) return null;
    const w = Math.max(40, contentWidth - FN_INDENT);
    let height = 0;
    const out = paras.map((p) => {
      const lines = getLines(p, w);
      height += p.style.spaceBeforePx + totalHeight(lines) + p.style.spaceAfterPx;
      return { p, lines };
    });
    return { paras: out, height };
  };
  /** New (unplaced) refs inside lines[i..i+take): [refId, markerText][] */
  const refsInLines = (lines: LineBox[], i: number, take: number): [string, string][] => {
    const out: [string, string][] = [];
    for (let k = i; k < i + take; k++) {
      for (const f of lines[k]!.fragments) {
        const id = f.style.footnoteRef;
        if (id && fns[id] && !placedNotes.has(id) && !out.some(([o]) => o === id)) out.push([id, f.text]);
      }
    }
    return out;
  };

  // Column cursor: blocks place at colX(); flow-driven breaks advance the
  // column first, then the page. "Empty page" guards become "empty COLUMN"
  // guards (forced progress, atomic placement) — tracked by where the current
  // column started in page.blocks.
  let colIdx = 0;
  let colStartCount = 0;
  const colX = (): number => colBoxes[colIdx]?.x ?? contentX;
  const colHasContent = (): boolean => page.blocks.length > colStartCount;

  /** Explicit page break / section start: always a fresh page, column 1. */
  const hardPage = (): void => {
    colIdx = 0;
    colStartCount = 0;
    page = mkPage();
    pages.push(page);
    pageNotes.push([]);
    pageReserved.push(0);
    reserved = 0; // footnote reservations are per PAGE (columns share the area)
    y = contentTop; // note: space-before is suppressed at the top of a page (Word behavior)
    floats = []; // floats never cross pages (or columns)
  };

  /** Flow advance: next column, or next page when columns are exhausted. */
  const newPage = (): void => {
    if (colIdx + 1 < colCount) {
      colIdx++;
      colStartCount = page.blocks.length;
      y = contentTop;
      floats = [];
    } else {
      hardPage();
    }
  };

  /** Re-stamp the current (empty) page with the active section's geometry instead
   *  of leaving a blank page when a section starts at a fresh page boundary. */
  const restampEmptyPage = (): void => {
    page.widthPx = sec.pageWidthPx;
    page.heightPx = sec.pageHeightPx;
    page.marginPx = sec.marginPx;
    page.contentTopPx = contentTop;
    page.contentBottomPx = contentBottom;
    if (sec.pageColorHex !== undefined) page.pageColorHex = sec.pageColorHex;
    else delete page.pageColorHex;
    if (sec.pageBorders !== undefined) page.pageBorders = sec.pageBorders;
    else delete page.pageBorders;
    if (colSeparatorsX.length) page.columnSeparatorsX = colSeparatorsX.slice();
    else delete page.columnSeparatorsX;
    pageSections[page.index] = sec;
    colIdx = 0;
    colStartCount = 0;
    y = contentTop;
    floats = [];
  };

  /** Section boundary: start the next section's first page. An EMPTY current
   *  page is re-stamped with the new geometry instead of leaving a blank page.
   *  For an evenPage/oddPage break the section's first content page must land on
   *  an even/odd page number — a blank filler page is inserted when the running
   *  parity is wrong (Word's behavior). Parity is read off the provisional 1-based
   *  page number, which equals the printed number absent a pageNumberStart restart. */
  const startSectionPage = (breakType: SectionBreakType): void => {
    if (page.blocks.length === 0) restampEmptyPage();
    else hardPage();
    if (breakType === "evenPage" || breakType === "oddPage") {
      const wantEven = breakType === "evenPage";
      if ((page.number % 2 === 0) !== wantEven) {
        // Current page is the wrong parity: leave it blank (the filler) and push a
        // fresh page (hardPage stamps it from the active section) for the content.
        hardPage();
      }
    }
  };

  /** Absolute x of a list marker. LTR: it hangs `hangingPx` left of the text
   *  start (placedX). RTL: the indent gutter is on the RIGHT, so the marker hangs
   *  just past the line's (right-aligned) right edge, pushed to the far side of the
   *  gutter to mirror the LTR gap. */
  const markerX = (
    block: Paragraph,
    placedX: number,
    line: LineBox | undefined,
    marker: { text: string; style: CharStyle; hangingPx: number },
  ): number => {
    if (block.style.direction !== "rtl") return placedX - marker.hangingPx;
    let rightEdge = 0;
    for (const f of line?.fragments ?? []) rightEdge = Math.max(rightEdge, f.x + f.width);
    const markerW = measureTextWidth(marker.text, charStyleToFont(marker.style));
    return placedX + rightEdge + Math.max(0, marker.hangingPx - markerW);
  };

  const placeRun = (block: Paragraph, lines: LineBox[], from: number, count: number): void => {
    const placed: PlacedBlock = {
      blockId: block.id,
      x: colX() + indentOf(block),
      y,
      firstLineIndex: from,
      lines: [],
    };
    const marker = from === 0 ? markers.get(block.id) : undefined;
    if (marker) {
      placed.marker = { text: marker.text, style: marker.style, x: markerX(block, placed.x, lines[from], marker) };
    }
    if (block.style.tocEntry) {
      // The number decorates the LAST chunk's last line; right edge = this column's.
      tocPlacements.set(block.id, { chunk: placed, rightEdge: colX() + colWidth, para: block });
    }
    page.blocks.push(placed);
    let chunkHeight = 0;
    for (let k = from; k < from + count; k++) {
      const line = lines[k]!;
      placed.lines.push({ ...line, y: y - placed.y });
      y += line.height;
      chunkHeight += line.height;
    }
    // Paragraph shading / border box (w:shd / w:pBdr): the box spans this column's
    // content edges between the paragraph's indents, and this chunk's lines tall.
    // Painted behind the text in renderer.ts / pdf paintBlock.ts.
    const decor = paraDecorFor(block.style, colWidth - indentOf(block) - rightIndentOf(block), chunkHeight);
    if (decor) placed.paraDecor = decor;
  };

  /** How many lines starting at `from` fit above the effective bottom at the current y. */
  const countFit = (lines: LineBox[], from: number): number => {
    let fit = 0;
    let yy = y;
    while (from + fit < lines.length && yy + lines[from + fit]!.height <= bottomY()) {
      yy += lines[from + fit]!.height;
      fit++;
    }
    return fit;
  };

  /** Page-bottom space the given refs' notes need (FN_SEP only when this page
   *  has no notes yet), paired with the measures to commit. */
  const measureNotes = (refs: [string, string][]): { H: number; measures: [string, string, NoteMeasure][] } => {
    let H = pageNotes[page.index]!.length === 0 ? FN_SEP : 0;
    const measures: [string, string, NoteMeasure][] = [];
    for (const [id, numText] of refs) {
      const m = noteMeasure(id);
      if (m) {
        H += m.height;
        measures.push([id, numText, m]);
      }
    }
    return { H, measures };
  };

  /** Reserve `H` px at this page's bottom and mark the notes placed. */
  const commitNotes = (measures: [string, string, NoteMeasure][], H: number): void => {
    for (const [id, numText] of measures) {
      placedNotes.add(id);
      pageNotes[page.index]!.push(id);
      noteNumbers.set(id, numText);
    }
    reserved += H;
    pageReserved[page.index] = reserved;
  };

  /** New (unplaced) footnote refs in a set of table rows (one level deep into
   *  nested tables), document order: [refId, markerText][]. */
  const refsInRows = (rows: TableBlock["rows"]): [string, string][] => {
    const out: [string, string][] = [];
    const visit = (blocks: Block[]): void => {
      for (const b of blocks) {
        if (b.kind === "paragraph") {
          let acc = 0;
          for (const r of b.runs) {
            const id = r.style.footnoteRef;
            if (id && fns[id] && !placedNotes.has(id) && !out.some(([o]) => o === id)) out.push([id, r.text]);
            acc += r.text.length;
          }
        } else if (b.kind === "table") {
          for (const row of b.rows) for (const cell of row.cells) visit(cell.blocks);
        }
      }
    };
    for (const row of rows) for (const cell of row.cells) visit(cell.blocks);
    return out;
  };

  /** Reserve footnote space for the new refs in lines[i..i+take), shrinking
   *  `take` until the chunk and its notes co-fit on this page. Mutates the
   *  reservation state when it commits. */
  const fitChunkWithNotes = (lines: LineBox[], i: number, take: number): number => {
    for (;;) {
      if (take <= 0) return take;
      const refs = refsInLines(lines, i, take);
      if (refs.length === 0) return take;
      const { H, measures } = measureNotes(refs);
      if (measures.length === 0) return take;
      let chunkH = 0;
      for (let k = i; k < i + take; k++) chunkH += lines[k]!.height;
      const fits = y + chunkH + H <= bottomY();
      if (fits || (take === 1 && !colHasContent())) {
        // Commit (overflow rather than loop when even one line can't co-fit
        // in an empty column — the note may then collide; Word splits notes,
        // we don't yet).
        commitNotes(measures, H);
        return take;
      }
      take--;
    }
  };

  const placeParagraph = (block: Paragraph, lines: LineBox[]): void => {
    // keep-lines-together: a paragraph that would split moves WHOLE to the
    // next column/page — unless it can't fit on an empty one (then it splits
    // normally, like Word).
    if (block.style.keepLinesTogether === true && colHasContent()) {
      const total = totalHeight(lines);
      if (y + total > bottomY() && total <= contentBottom - contentTop) newPage();
    }
    let i = 0;
    while (i < lines.length) {
      const remaining = lines.length - i;
      const fit = countFit(lines, i);

      let take: number;
      if (fit >= remaining) {
        take = remaining;
      } else {
        take = fit;
        // Widow/orphan control is Word's default (ON); w:widowControl="0" disables it,
        // letting a lone first/last line break across the page boundary.
        const widow = block.style.widowControl !== false;
        // Widow: never push a single last line to the next page — give it company.
        if (widow && remaining - take === 1) take -= WIDOW_MIN - 1;
        // Orphan: at the paragraph's start, keep ≥2 lines here or move it whole.
        if (widow && i === 0 && take > 0 && take < ORPHAN_MIN) take = 0;
        // Forced progress: in an empty column the rules yield (a line taller than
        // the page, or a 3-line paragraph that can't satisfy 2+2) — place what fits.
        if (take <= 0) take = !colHasContent() ? Math.max(1, fit) : 0;
      }

      // Footnotes referenced by the chunk reserve page-bottom space; the chunk
      // shrinks until both co-fit (a smaller chunk references fewer notes).
      if (take > 0 && hasNotes) take = fitChunkWithNotes(lines, i, take);

      if (take > 0) {
        placeRun(block, lines, i, take);
        i += take;
      }
      if (i >= lines.length) return;
      newPage();
    }
  };

  const totalHeight = (lines: LineBox[]): number => {
    let h = 0;
    for (const l of lines) h += l.height;
    return h;
  };

  // Active floats on the CURRENT page (wrap:'square' images). Cleared on page
  // break — floats never cross pages.
  interface Float {
    left: number;
    right: number;
    bottom: number;
    side: "left" | "right";
  }
  let floats: Float[] = [];
  const FLOAT_GUTTER = 10;

  /** Line box available at vertical position yy within the CURRENT column,
   *  shrunk by active floats. */
  const lineBoxAt = (yy: number): { x0: number; width: number } => {
    let x0 = colX();
    let x1 = colX() + colWidth;
    for (const f of floats) {
      if (yy >= f.bottom) continue;
      if (f.side === "left") x0 = Math.max(x0, f.right + FLOAT_GUTTER);
      else x1 = Math.min(x1, f.left - FLOAT_GUTTER);
    }
    return { x0, width: Math.max(40, x1 - x0) };
  };

  const floatsActiveAt = (yy: number): boolean => floats.some((f) => f.bottom > yy);

  /** wp:anchor + wp:wrapNone: an absolutely-positioned behind/in-front image.
   *  Positioned at its offset from the relativeFrom origin; it does NOT advance
   *  the flow cursor and registers NO float (text flows OVER it, not around).
   *  Z-order falls out of document order — a background image precedes the body
   *  in the block list, so it paints first (behind). */
  const placeAnchoredImage = (img: ImageBlock): void => {
    const a = img.anchor!;
    let ox: number;
    switch (a.relFromH) {
      case "page":
      case "leftMargin": ox = 0; break;
      case "rightMargin": ox = sec.pageWidthPx - sec.marginPx.right; break;
      case "column": ox = colX(); break;
      default: ox = sec.marginPx.left; break; // margin / character
    }
    let oy: number;
    switch (a.relFromV) {
      case "page":
      case "topMargin": oy = 0; break;
      case "bottomMargin": oy = sec.pageHeightPx - sec.marginPx.bottom; break;
      case "paragraph":
      case "line": oy = y; break; // current flow position
      default: oy = sec.marginPx.top; break; // margin
    }
    const placedImage: PlacedImage = { src: img.src, width: img.widthPx, height: img.heightPx, z: a.z ?? 0, ...(img.crop ? { crop: img.crop } : {}) };
    if (a.behind) placedImage.behind = true;
    else placedImage.front = true;
    page.blocks.push({
      blockId: img.id,
      x: ox + a.offsetXPx,
      y: oy + a.offsetYPx,
      firstLineIndex: 0,
      lines: [],
      image: placedImage,
    });
  };

  const placeImage = (img: ImageBlock): void => {
    if (img.anchor) { placeAnchoredImage(img); return; }
    if (y + img.heightPx > bottomY() && colHasContent()) newPage();
    const floating = img.wrap === "square" && img.align !== "center";
    const slack = colWidth - img.widthPx;
    const x =
      colX() + (img.align === "center" ? slack / 2 : img.align === "right" ? slack : 0);
    page.blocks.push({
      blockId: img.id,
      x,
      y,
      firstLineIndex: 0,
      lines: [],
      image: { src: img.src, width: img.widthPx, height: img.heightPx, ...(img.crop ? { crop: img.crop } : {}) },
    });
    if (floating) {
      // Text flows beside the image: register the float, do NOT advance y.
      floats.push({ left: x, right: x + img.widthPx, bottom: y + img.heightPx, side: img.align === "right" ? "right" : "left" });
    } else {
      y += img.heightPx;
    }
  };

  const placeEquation = (eqb: EquationBlock, box: MathBox): void => {
    const h = box.ascent + box.descent;
    if (y + h > bottomY() && colHasContent()) newPage();
    const align = eqb.align ?? "center";
    const slack = colWidth - box.width;
    const x = colX() + (align === "center" ? slack / 2 : align === "right" ? slack : 0);
    page.blocks.push({
      blockId: eqb.id,
      x,
      y,
      firstLineIndex: 0,
      lines: [],
      equation: { box, width: box.width, height: h, baseline: box.ascent },
    });
    y += h;
  };

  /** Float-affected paragraphs: re-break per line with the float-shrunk width
   *  at the line's own y (pretext's per-line maxWidth makes this cheap). A line
   *  that doesn't fit the page is ROLLED BACK (runCursors snapshot) and
   *  re-broken at full width on the next page. Justify is applied with a
   *  one-line buffer (a line is justified once the NEXT line proves it isn't the
   *  segment's last — which stays ragged, like a paragraph's final line). */
  const placeParagraphFloating = (p: Paragraph): void => {
    const segments = getPrepared(p);
    const levels = paragraphLevels(p);
    const palign = effectiveAlign(p.style.align, p.style.direction);
    let first = true;
    let lineIdx = 0;
    let placed: PlacedBlock | null = null;

    // Attach the paragraph's shading/border box to the chunk currently being
    // built — once before each page-break reset and once at the end — so a boxed
    // paragraph beside a float still renders its decoration (per chunk).
    const decorateFloat = (): void => {
      if (placed && placed.lines.length) {
        const d = paraDecorFor(p.style, colWidth - indentOf(p) - rightIndentOf(p), sumLineHeights(placed.lines));
        if (d) placed.paraDecor = d;
      }
    };

    const pushLine = (
      frags: RawFrag[],
      height: number,
      ascent: number,
      emptyOffset?: number,
    ): void => {
      if (placed === null) {
        placed = { blockId: p.id, x: colX() + indentOf(p), y, firstLineIndex: lineIdx, lines: [] };
        const marker = lineIdx === 0 ? markers.get(p.id) : undefined;
        if (marker) {
          // Beside a float the line text is shifted right (frag.x carries the float
          // offset); hang the marker off the FIRST line's actual start so it tracks
          // the text instead of stranding at the left margin ON TOP of the float.
          // Equals placed.x - hangingPx when there's no shift (frag.x === 0).
          // RTL: hang past the line's right edge instead (mirrored gutter).
          let mx: number;
          if (p.style.direction === "rtl") {
            let rightEdge = 0;
            for (const rf of frags) rightEdge = Math.max(rightEdge, rf.frag.x + rf.frag.width);
            const markerW = measureTextWidth(marker.text, charStyleToFont(marker.style));
            mx = placed.x + rightEdge + Math.max(0, marker.hangingPx - markerW);
          } else {
            mx = placed.x + (frags[0]?.frag.x ?? 0) - marker.hangingPx;
          }
          placed.marker = { text: marker.text, style: marker.style, x: mx };
        }
        page.blocks.push(placed);
      }
      const box: LineBox = { y: y - placed.y, height, ascent, fragments: frags.map((rf) => rf.frag) };
      if (emptyOffset !== undefined) box.emptyOffset = emptyOffset;
      placed.lines.push(box);
      y += height;
      lineIdx++;
    };

    for (const seg of segments) {
      const runStarts = segRunOffsets(seg.runs, seg.startOffset);
      const runCursors: number[] = seg.runs.map(() => 0);
      let cursor: RichInlineCursor | undefined = undefined;
      let produced = false;
      // Buffered previous line for justify: held until the next line confirms it
      // isn't the segment's last (the last line stays ragged). Reset per segment.
      let justifyPrev: { frags: RawFrag[]; slack: number } | null = null;

      for (;;) {
        const snapshot = runCursors.slice();
        const box = lineBoxAt(y);
        const indent = first ? p.style.indentFirstLinePx : 0;
        const bl = breakNextLine(
          p,
          seg,
          Math.max(24, box.width - indentOf(p) - indent - rightIndentOf(p)),
          cursor,
          runStarts,
          runCursors,
        );
        if (bl === null) break;
        if (y + bl.height > bottomY() && colHasContent()) {
          for (let k = 0; k < runCursors.length; k++) runCursors[k] = snapshot[k]!;
          decorateFloat();
          newPage(); // floats cleared; the line re-breaks at full width
          placed = null;
          continue;
        }
        if (levels !== null && bl.frags.length > 0) {
          const r = bidiReorderLine(bl.frags, levels);
          bl.frags = r.frags;
          bl.width = r.width;
        }
        const slack = box.width - indent - rightIndentOf(p) - bl.width;
        let startX = box.x0 - colX() + indent;
        if (palign === "center") startX += slack / 2;
        else if (palign === "right") startX += slack;
        for (const rf of bl.frags) rf.frag.x += startX;
        // Justify the PREVIOUS line now that this one exists (so it's non-last);
        // buffer this line for the same decision next iteration. Justify runs on the
        // visual fragment order, so bidi-reordered lines justify too.
        if (palign === "justify") {
          if (justifyPrev) justifyLine(justifyPrev.frags, justifyPrev.slack);
          justifyPrev = { frags: bl.frags, slack };
        }
        pushLine(bl.frags, bl.height, bl.ascent);
        cursor = bl.end;
        first = false;
        produced = true;
      }

      if (!produced) {
        const m = emptyLineMetrics(p);
        if (y + m.height > bottomY() && colHasContent()) {
          decorateFloat();
          newPage();
          placed = null;
        }
        pushLine([], m.height, m.ascent, seg.startOffset);
        first = false;
      }
    }
    decorateFloat();
  };

  /** Tables split at ROW boundaries across pages (Word behavior). At least one
   *  row per chunk; a lone row taller than the page overflows in place. */
  const placeTableChunked = (m: Measured & { kind: "table" }): void => {
    // A vertical merge can't survive a row-boundary split — the rowSpan cell (and
    // usually the header row it shares) live only in the first chunk, leaving the
    // continuation page with an empty column. So a merged table that fits on a
    // full page moves WHOLE to the next column/page rather than splitting (Word's
    // effective behavior for these comparable grids). Taller-than-a-page merged
    // tables still split (unavoidable).
    const hasRowSpan = m.block.rows.some((r) => r.cells.some((c) => (c.rowSpan ?? 1) > 1));
    if (hasRowSpan && colHasContent() && y + m.height > bottomY() && m.height <= contentBottom - contentTop) {
      newPage();
    }
    // Repeat-header rows (w:tblHeader): the LEADING contiguous header rows are
    // re-drawn at the top of every continuation chunk. They appear normally inside
    // the first chunk (rows 0..); on later pages we place them as a SEPARATE table
    // block above the data chunk, which keeps each placed block's rows contiguous
    // (the invariant geometry/cellRangeRects rely on). Disabled when a vertical
    // merge is present (a rowSpan can't survive being detached from its band) or
    // when every row is a header (nothing left to continue).
    let headerCount = 0;
    while (headerCount < m.rows.length && m.block.rows[headerCount]?.props?.repeatHeader) headerCount++;
    const repeatHeader = headerCount > 0 && headerCount < m.rows.length && !hasRowSpan;
    const headerRows = repeatHeader ? m.rows.slice(0, headerCount) : [];
    const headerHeight = headerRows.reduce((s, r) => s + r.height, 0);
    // w:cantSplit needs no dedicated branch here: this paginator is ROW-ATOMIC — the
    // chunk loop below only ever slices at row boundaries (m.rows.slice(ri, ri+fit))
    // and places whole measured rows, so a row's cells are never broken across a page.
    // Every row is therefore kept-whole already; `props.cantSplit` round-trips for
    // fidelity and stays the marker to exclude a row should true intra-row splitting
    // ever be added. (A row taller than a full empty column still overflows — there is
    // nowhere to move it — matching the no-cantSplit case.)
    let ri = 0;
    while (ri < m.rows.length) {
      // Once we are past the original header band, every continuation chunk reserves
      // space for the repeated header at the top.
      const repeat = repeatHeader && ri >= headerCount;
      const topUsed = repeat ? headerHeight : 0;
      let fit = 0;
      let yy = y + topUsed;
      while (ri + fit < m.rows.length && yy + m.rows[ri + fit]!.height <= bottomY()) {
        yy += m.rows[ri + fit]!.height;
        fit++;
      }
      if (fit === 0) {
        if (colHasContent()) {
          newPage();
          continue;
        }
        fit = 1; // row taller than an empty column: overflow rather than loop
      }
      // Footnote refs inside these rows reserve page-bottom space; shrink the
      // chunk until rows + notes co-fit (the table analogue of fitChunkWithNotes).
      let pendingNotes: [string, string, NoteMeasure][] = [];
      let pendingNotesH = 0;
      for (;;) {
        const refs = refsInRows(m.block.rows.slice(ri, ri + fit));
        if (refs.length === 0) break;
        const { H, measures } = measureNotes(refs);
        if (measures.length === 0) break;
        let chunkH = topUsed;
        for (let k = ri; k < ri + fit; k++) chunkH += m.rows[k]!.height;
        if (y + chunkH + H <= bottomY() || (fit === 1 && !colHasContent())) {
          pendingNotes = measures;
          pendingNotesH = H;
          break;
        }
        fit--;
        if (fit === 0) break; // even one row + its notes overflow → next page
      }
      if (fit === 0) {
        newPage();
        continue;
      }
      if (pendingNotes.length > 0) commitNotes(pendingNotes, pendingNotesH);
      if (repeat) {
        // Repeated header band — its own contiguous placed block at logical rows 0..
        page.blocks.push(placeTable(m.block, headerRows, m.colWidths, colX() + m.xOffset, y, m.tableWidth, 0, cellListCtx));
        y += headerHeight;
      }
      const chunk = m.rows.slice(ri, ri + fit);
      page.blocks.push(placeTable(m.block, chunk, m.colWidths, colX() + m.xOffset, y, m.tableWidth, ri, cellListCtx));
      y += chunk.reduce((s, r) => s + r.height, 0);
      ri += fit;
      if (ri < m.rows.length) newPage();
    }
  };

  // Contextual spacing (docx w:contextualSpacing): a paragraph flagged with it
  // drops its before/after spacing against an adjacent SAME-STYLE paragraph
  // (Word's list-style default), so a run of same-style paragraphs sits tight
  // while the run's outer edges keep their spacing. True when `cur` is such a
  // paragraph sitting flush against same-named-style paragraph `adj`.
  const csSuppresses = (cur: Measured | undefined, adj: Measured | undefined): boolean =>
    cur?.kind === "para" &&
    adj?.kind === "para" &&
    cur.block.style.contextualSpacing === true &&
    cur.block.style.namedStyle === adj.block.style.namedStyle;

  for (let bi = 0; bi < measured.length; bi++) {
    // Crossing into a new section: swap geometry, force its first page.
    if (bi > sections[secIdx]!.endBlock) {
      while (bi > sections[secIdx]!.endBlock) applySection(sections[++secIdx]!.props);
      startSectionPage(sections[secIdx]!.breakType);
    }
    const m = measured[bi]!;

    // Two directly-adjacent atomics (a title-bar table stacked on its data grid,
    // back-to-back images) sit FLUSH — the gap belongs between an atomic and
    // surrounding text, not between atomics, so don't double it into an empty line.
    // Anchored (behind/in-front) images are out of flow — transparent to the
    // gap logic so they neither take a gap nor suppress one between neighbours.
    const atomicKind = (x: Measured | undefined): boolean =>
      x?.kind === "table" || x?.kind === "equation" || (x?.kind === "image" && x.block.anchor === undefined);
    const prevAtomic = atomicKind(measured[bi - 1]);
    const nextAtomic = atomicKind(measured[bi + 1]);

    if (m.kind === "image") {
      // Out-of-flow anchor: place absolutely, advance nothing, no gap.
      if (m.block.anchor) { placeImage(m.block); continue; }
      if (!prevAtomic) y += ATOMIC_GAP;
      placeImage(m.block);
      if (!nextAtomic && (m.block.wrap !== "square" || m.block.align === "center")) y += ATOMIC_GAP;
      continue;
    }
    if (m.kind === "equation") {
      if (!prevAtomic) y += ATOMIC_GAP;
      for (const f of floats) if (f.bottom > y) y = f.bottom;
      placeEquation(m.block, m.box);
      if (!nextAtomic) y += ATOMIC_GAP;
      continue;
    }
    if (m.kind === "table") {
      if (!prevAtomic) y += ATOMIC_GAP;
      // Tables don't flow beside floats the way text does — drop below any float
      // still active at y so the grid never overlaps a floated image.
      for (const f of floats) if (f.bottom > y) y = f.bottom;
      placeTableChunked(m);
      if (!nextAtomic) y += ATOMIC_GAP;
      continue;
    }

    const { block, lines } = m;
    // Page break = fresh PAGE (even from column 2); column break = next column.
    if (block.style.pageBreakBefore === true && (page.blocks.length > 0 || colIdx > 0)) hardPage();
    else if (block.style.columnBreakBefore === true && colHasContent()) newPage();
    // Contextual spacing: suppress this paragraph's outer spacing against an
    // adjacent same-style paragraph (see csSuppresses).
    const suppressBefore = csSuppresses(m, measured[bi - 1]);
    const suppressAfter = csSuppresses(m, measured[bi + 1]);
    if (!suppressBefore) y += block.style.spaceBeforePx;
    if (y >= bottomY() && colHasContent()) newPage();

    // Float-affected paragraphs re-break per line with float-shrunk widths —
    // bypassing the precomputed full-width lines (and widow/orphan rules).
    if (floatsActiveAt(y)) {
      // Reserve the border box (rule widths + pad) around the float-adjacent
      // chunk too — placeParagraphFloating still attaches paraDecor with `pad`,
      // so the box would otherwise extend into the next block (w:pBdr).
      y += paraBorderReserveTop(block.style);
      placeParagraphFloating(block);
      y += paraBorderReserveBottom(block.style);
      if (!suppressAfter) y += block.style.spaceAfterPx;
      continue;
    }

    // Keep-with-next: if this block fits here but the keep-CHAIN hanging off it
    // can't start on this page (each keepWithNext member must fit WHOLE, and
    // the first non-keep successor must keep ≥ its orphan/widow take —
    // mirroring placeParagraph's rules), break before this block instead.
    if (block.style.keepWithNext === true && measured[bi + 1] !== undefined && colHasContent()) {
      // Mirror the contextual-spacing suppression the real placement applies, so
      // a same-style contextual run isn't over-measured and broken when it fits.
      let yy = y + totalHeight(lines) + (suppressAfter ? 0 : block.style.spaceAfterPx);
      let ok = yy <= bottomY();
      for (let k = bi + 1; ok; k++) {
        const m2 = measured[k];
        if (m2 === undefined) break;
        if (m2.kind !== "para") {
          // Atomic chain end: images need to fit entirely; tables only their
          // first row (they chunk at row boundaries).
          const firstUnit = m2.kind === "table" ? (m2.rows[0]?.height ?? m2.height) : m2.height;
          if (yy + ATOMIC_GAP + firstUnit > bottomY()) ok = false;
          break;
        }
        if (!csSuppresses(m2, measured[k - 1])) yy += m2.block.style.spaceBeforePx;
        const afterGap = csSuppresses(m2, measured[k + 1]) ? 0 : m2.block.style.spaceAfterPx;
        const inChain = m2.block.style.keepWithNext === true && measured[k + 1] !== undefined;
        if (inChain || m2.block.style.keepLinesTogether === true) {
          // Chain member (or keep-lines paragraph): must fit WHOLE.
          const h = totalHeight(m2.lines);
          if (yy + h > bottomY()) {
            ok = false;
            break;
          }
          yy += h + afterGap;
          if (!inChain) break;
          continue;
        }
        // An empty spacer paragraph (e.g. a blank line between a "keep with next"
        // heading and its table) must NOT satisfy the keep — otherwise the heading
        // sticks to the blank line and orphans from the content it belongs with.
        // Bridge through it (fit it whole) to the real next block.
        if (measured[k + 1] !== undefined && m2.block.runs.every((r) => r.text.trim().length === 0)) {
          const h = totalHeight(m2.lines);
          if (yy + h > bottomY()) {
            ok = false;
            break;
          }
          yy += h + afterGap;
          continue;
        }
        // Chain terminator: needs its orphan/widow-legal first take.
        let fit = 0;
        let fy = yy;
        while (fit < m2.lines.length && fy + m2.lines[fit]!.height <= bottomY()) {
          fy += m2.lines[fit]!.height;
          fit++;
        }
        let take = fit;
        // Apply the orphan/widow take only when the terminator has widow control on
        // (default); widowControl:false lets a lone line stand, so don't force a page.
        if (fit < m2.lines.length && m2.block.style.widowControl !== false) {
          if (m2.lines.length - take === 1) take -= WIDOW_MIN - 1; // widow
          if (take < ORPHAN_MIN) take = 0; // orphan
        }
        if (take <= 0) ok = false;
        break;
      }
      if (!ok) newPage();
    }

    // Reserve the paragraph border box (top/bottom rule widths + border-to-text
    // padding) in the surrounding gaps so a boxed paragraph's box — expanded
    // outward by `pad`, and any shading fill — doesn't paint over the adjacent
    // block (w:pBdr).
    y += paraBorderReserveTop(block.style);
    placeParagraph(block, lines);
    y += paraBorderReserveBottom(block.style);
    if (!suppressAfter) y += block.style.spaceAfterPx;
  }

  // Margin line numbers (w:lnNumType): walk pages in order and number each body
  // text line of every line-numbered section. Runs BEFORE footnote paragraphs are
  // appended below, so notes are never numbered, and reads page.blocks in flow
  // order (which is vertical order within a column). `countBy` selects which
  // numbers show; `restart` controls when the counter resets — "newPage" (Word's
  // default), "newSection", or "continuous" (straight through). Numbers are
  // paint-only, pre-measured and right-aligned into the left margin so both the
  // canvas renderer and PDF painter just fillText at (x, baseline).
  if (pageSections.some((s) => s.lineNumbering)) {
    const font = charStyleToFont(LINE_NUMBER_STYLE);
    let counter = 1;
    let started = false; // "continuous": initialise the counter only once
    let prevSection: SectionProps | null = null;
    for (const pg of pages) {
      const s = pageSections[pg.index]!;
      const ln: LineNumbering | undefined = s.lineNumbering;
      if (!ln) {
        prevSection = s;
        continue;
      }
      const start = ln.start ?? 1;
      const countBy = Math.max(1, Math.round(ln.countBy ?? 1));
      const restart = ln.restart ?? "newPage";
      if (restart === "newPage") counter = start;
      else if (restart === "newSection") {
        if (s !== prevSection) counter = start;
      } else if (!started) counter = start; // continuous
      started = true;
      const rightEdge = pg.marginPx.left - (ln.distancePx ?? LINE_NUMBER_DISTANCE);
      const labels: NonNullable<Page["lineNumbers"]> = [];
      for (const b of pg.blocks) {
        if (b.lines.length === 0 || b.image || b.table || b.equation) continue; // body paragraphs only
        for (const line of b.lines) {
          if (counter % countBy === 0) {
            const text = String(counter);
            labels.push({
              x: rightEdge - measureTextWidth(text, font),
              baseline: b.y + line.y + line.ascent,
              text,
              style: LINE_NUMBER_STYLE,
            });
          }
          counter++;
        }
      }
      if (labels.length > 0) pg.lineNumbers = labels;
      prevSection = s;
    }
  }

  // Footnote areas: each page's notes stack at the bottom under a separator
  // rule, inside the space reserved during the walk. Note paragraphs are REAL
  // model paragraphs pushed as page blocks — geometry indexes them, so caret,
  // selection, and typing inside notes work with zero special cases. The
  // number marker (the ref run's text) is paint-only, hung in the indent.
  for (const pg of pages) {
    const ids = pageNotes[pg.index];
    if (!ids || ids.length === 0) continue;
    const ps = pageSections[pg.index]!;
    applySection(ps); // contentWidth/contentBottom for noteMeasure + placement
    let ny = contentBottom - pageReserved[pg.index]!;
    pg.footnoteRuleY = ny + 4;
    ny += FN_SEP;
    const nx = ps.marginPx.left;
    for (const id of ids) {
      const m = noteMeasure(id);
      if (!m) continue;
      let first = true;
      for (const np of m.paras) {
        ny += np.p.style.spaceBeforePx;
        const placed: PlacedBlock = {
          blockId: np.p.id,
          x: nx + FN_INDENT,
          y: ny,
          firstLineIndex: 0,
          lines: np.lines,
        };
        if (first && np.p.runs[0]) {
          placed.marker = { text: `${noteNumbers.get(id) ?? "•"}.`, style: np.p.runs[0].style, x: nx };
          first = false;
        }
        const noteDecor = paraDecorFor(
          np.p.style,
          contentWidth - FN_INDENT - np.p.style.indentLeftPx - (np.p.style.indentRightPx ?? 0),
          totalHeight(np.lines),
        );
        if (noteDecor) placed.paraDecor = noteDecor;
        pg.blocks.push(placed);
        ny += totalHeight(np.lines) + np.p.style.spaceAfterPx;
      }
    }
  }

  // Endnotes: collected at the END of the document, under a separator rule, in
  // reference order (Word's "end of document" placement — the counterpart to the
  // page-bottom footnote area above). Like footnotes, note paragraphs are REAL
  // model paragraphs pushed as page blocks so caret/selection work; the number
  // marker (the ref run's text) is paint-only, hung in the indent. Notes flow
  // onto continuation pages but aren't split across them (matching footnotes).
  const ens = doc.endnotes ?? {};
  if (Object.keys(ens).length > 0) {
    // References in document order → [noteId, markerText], deduped.
    const enRefs: [string, string][] = [];
    const seenEn = new Set<string>();
    const collectEn = (blocks: Block[]): void => {
      for (const b of blocks) {
        if (b.kind === "paragraph") {
          for (const r of b.runs) {
            const id = r.style.endnoteRef;
            if (id && ens[id] && !seenEn.has(id)) {
              seenEn.add(id);
              enRefs.push([id, r.text]);
            }
          }
        } else if (b.kind === "table") {
          for (const row of b.rows) for (const cell of row.cells) collectEn(cell.blocks);
        }
      }
    };
    collectEn(doc.blocks);
    if (enRefs.length > 0) {
      // Endnotes use the document-end (last) section geometry, single column.
      applySection(sections[sections.length - 1]!.props);
      const enWidth = Math.max(40, contentWidth - FN_INDENT);
      const enx = sec.marginPx.left;
      let epage = pages[pages.length - 1]!;
      const newEndnotePage = (): void => {
        epage = mkPage();
        delete epage.columnSeparatorsX; // endnotes are a full-width single-column story
        pages.push(epage);
        pageNotes.push([]);
        pageReserved.push(0);
      };
      // Start on a fresh page if the last body page already has content.
      if (epage.blocks.length > 0) newEndnotePage();
      else delete epage.columnSeparatorsX; // reused last page → drop its column rules too
      let ey = epage.contentTopPx;
      epage.endnoteRuleY = ey + 4;
      ey += FN_SEP;
      for (const [id, numText] of enRefs) {
        let first = true;
        for (const p of ens[id]!) {
          const lines = getLines(p, enWidth);
          // Page-break before a note paragraph that won't fit (notes aren't split).
          // Gate on the page already holding content — never bump the first note
          // off a fresh page (whose `ey` is past the top only because of the rule),
          // which would strand a blank separator-only page.
          if (ey + p.style.spaceBeforePx + totalHeight(lines) > epage.contentBottomPx && epage.blocks.length > 0) {
            newEndnotePage();
            ey = epage.contentTopPx;
          }
          ey += p.style.spaceBeforePx;
          const placed: PlacedBlock = {
            blockId: p.id,
            x: enx + FN_INDENT,
            y: ey,
            firstLineIndex: 0,
            lines,
          };
          if (first && p.runs[0]) {
            placed.marker = { text: `${numText}.`, style: p.runs[0].style, x: enx };
            first = false;
          }
          const noteDecor = paraDecorFor(
            p.style,
            contentWidth - FN_INDENT - p.style.indentLeftPx - (p.style.indentRightPx ?? 0),
            totalHeight(lines),
          );
          if (noteDecor) placed.paraDecor = noteDecor;
          epage.blocks.push(placed);
          ey += totalHeight(lines) + p.style.spaceAfterPx;
        }
      }
    }
  }

  // Z-order: anchored images are lifted out of the flow into a behind-text or
  // in-front-of-text layer. Both the canvas renderer and the PDF painter draw
  // blocks in array order, so the FINAL paint order IS this array order —
  // making this the single source of truth means the two renderers can never
  // drift. Block lists with no anchored images are left untouched (byte-identical
  // output for every ordinary document — no PDF drift risk). Recurses into table
  // cells (and nested tables) so behind/front layering works inside cells too.
  const reorderAnchoredLayers = (blocks: PlacedBlock[]): PlacedBlock[] => {
    for (const b of blocks) {
      if (b.table) for (const row of b.table.rows) for (const cell of row.cells) cell.blocks = reorderAnchoredLayers(cell.blocks);
    }
    if (!blocks.some((b) => b.image?.behind || b.image?.front)) return blocks;
    // Stable layer partition (behind → flow/text → front); within an anchored
    // layer, higher z paints later (on top). Flow blocks keep document order.
    const layer = (b: PlacedBlock): number => (b.image?.behind ? 0 : b.image?.front ? 2 : 1);
    return blocks
      .map((b, i) => ({ b, i }))
      .sort((p, q) => {
        const dl = layer(p.b) - layer(q.b);
        if (dl !== 0) return dl;
        if (layer(p.b) !== 1) {
          const dz = (p.b.image!.z ?? 0) - (q.b.image!.z ?? 0);
          if (dz !== 0) return dz;
        }
        return p.i - q.i; // stable within a layer / equal z
      })
      .map((x) => x.b);
  };
  for (const pg of pages) pg.blocks = reorderAnchoredLayers(pg.blocks);

  // Displayed page numbers: continue counting across sections unless a
  // section restarts the sequence (Word's "start at"). {page} tokens and the
  // even/odd band variant pick both read THESE, not the physical index.
  const pageNumbers: number[] = [];
  {
    let n = 0;
    for (let i = 0; i < pages.length; i++) {
      const s = pageSections[i]!;
      const firstOfSection = i === 0 || pageSections[i - 1] !== s;
      if (firstOfSection && s.pageNumberStart !== undefined) n = s.pageNumberStart;
      else n++;
      pageNumbers.push(n);
      pages[i]!.number = n;
    }
  }

  // Inline body PAGE/NUMPAGES fields: their result run carries a {page}/{pages}
  // token (the marker that tells a real field from literal text). The body is
  // laid out ONCE, so — unlike per-page header/footer bands — these tokens were
  // never substituted and would paint literally ("page {page}"). Resolve them
  // here against the final page map, paint-only: the token run is isolated by its
  // fieldId so it's its own fragment — rewrite that fragment's text, re-measure
  // its (always shorter) width, map the shorter text back onto the model range via
  // offsetMap so caret/hit-testing stay exact, and slide the rest of the line left
  // to close the gap the wider token left. Resolution is clone-on-write per line
  // (resolveBlockTokens) so it never mutates the shared line/table caches — the
  // cached {page} token survives, so the number re-resolves every pass even when
  // the cache is reused (no relayout, no painter change). Untagged literal "{page}"
  // body text is left alone (fieldId-gated).
  const fieldDefs = doc.fields;
  if (fieldDefs) {
    const isPageNumField = (fid: string | undefined): boolean => {
      if (fid === undefined) return false;
      const def = fieldDefs[fid];
      const t = def && def.kind === "builtin" ? def.spec?.type : undefined;
      return t === "PAGE" || t === "NUMPAGES";
    };
    const resolveLineTokens = (line: LineBox, pageNum: number): void => {
      let dx = 0;
      for (const frag of line.fragments) {
        if (dx !== 0) frag.x += dx; // close the gap left by an earlier token on this line
        if (!isPageNumField(frag.style.fieldId)) continue;
        const sub = substituteTokens(frag.text, pageNum, pages.length);
        if (sub === frag.text) continue;
        const modelLen = frag.endOffset - frag.startOffset;
        const oldWidth = frag.width;
        frag.text = sub;
        frag.width = measureTextWidth(sub, charStyleToFont(frag.style));
        // Displayed text ({page} -> "3") is shorter than the model run: collapse
        // every displayed cluster onto the field's model range so a caret lands at
        // the field's start/end, never mid-token (reuses the pretext offsetMap path).
        frag.offsetMap = Array.from({ length: sub.length + 1 }, (_v, i) => (i === sub.length ? modelLen : 0));
        dx += frag.width - oldWidth;
      }
    };
    const lineHasPageField = (line: LineBox): boolean =>
      line.fragments.some((f) => isPageNumField(f.style.fieldId));
    const resolveBlockTokens = (b: PlacedBlock, pageNum: number): void => {
      // resolveLineTokens MUTATES fragment text/width/x/offsetMap. A placed block's
      // `lines` ALIAS cached LineBoxes (linesCache for body paragraphs, tableCache's
      // measured cells for table content), so resolving in place would bake the
      // resolved number INTO the cache. A later layout that reuses that cache
      // verbatim — e.g. a table pushed to a new page by a page break, its
      // revision+width unchanged — then keeps the stale number, because the cached
      // fragment text is already "1" (not "{page}") and substituteTokens no-ops.
      // (A column resize changes the width, busting the cache, which is why it
      // "fixed" the number — #48.) Clone-on-write: copy any line that carries a page
      // field, and its fragments, before resolving — the cached {page} token stays
      // pristine for the next pass. Lines with no page field are a no-op, so skip.
      let out: LineBox[] | null = null;
      for (let i = 0; i < b.lines.length; i++) {
        const line = b.lines[i]!;
        if (!lineHasPageField(line)) continue;
        const copy: LineBox = { ...line, fragments: line.fragments.map((f) => ({ ...f })) };
        resolveLineTokens(copy, pageNum);
        (out ??= b.lines.slice())[i] = copy;
      }
      if (out) b.lines = out;
      if (b.table) for (const row of b.table.rows) for (const cell of row.cells) for (const cb of cell.blocks) resolveBlockTokens(cb, pageNum);
    };
    for (const pg of pages) for (const b of pg.blocks) resolveBlockTokens(b, pageNumbers[pg.index]!);
  }

  // TOC page numbers: PAINT-ONLY decorations resolved against the final page
  // map — they can never go stale and never affect line breaking (entries are
  // measured with a reserved right gutter). One pass, no fixpoint.
  if (tocPlacements.size > 0) {
    const firstPageOf = new Map<string, number>(); // blockId -> displayed number
    const mapPage = (blocks: PlacedBlock[], num: number): void => {
      for (const b of blocks) {
        if (!firstPageOf.has(b.blockId)) firstPageOf.set(b.blockId, num);
        // Headings (TOC targets) can live inside table cells — recurse so their
        // page number resolves just like body paragraphs.
        if (b.table) for (const row of b.table.rows) for (const cell of row.cells) mapPage(cell.blocks, num);
      }
    };
    for (const pg of pages) mapPage(pg.blocks, pageNumbers[pg.index]!);
    for (const rec of tocPlacements.values()) {
      const target = rec.para.style.tocEntry!.targetId;
      const num = firstPageOf.get(target);
      if (num === undefined) continue; // dangling entry: heading was deleted
      const style = rec.para.runs[0]?.style;
      if (!style || rec.chunk.lines.length === 0) continue;
      const numText = String(num);
      rec.chunk.toc = {
        numText,
        numX: rec.rightEdge - measureTextWidth(numText, charStyleToFont(style)),
        lineIndex: rec.chunk.lines.length - 1,
        style,
        targetId: target,
      };
    }
  }

  // Margin-band stories: full blocks (paragraphs/images/tables) laid out per
  // page so {page}/{pages} tokens resolve. Band cache is per-relayout; the
  // substituted clones get per-page ids so prepares never collide. Each page
  // uses ITS section's bands and geometry (sections inherit doc.section bands
  // unless their patch overrides them); the First variant overrides on a
  // section's first page, the Even variant on even page NUMBERS (Word).
  const pickBand = (
    ps: SectionProps,
    kind: "header" | "footer",
    firstOfSection: boolean,
    even: boolean,
  ): { blocks: Block[]; source: (typeof BAND_KEYS)[number] } | null => {
    const first = kind === "header" ? ps.headerFirst : ps.footerFirst;
    const evenB = kind === "header" ? ps.headerEven : ps.footerEven;
    if (firstOfSection && first) return { blocks: first, source: kind === "header" ? "headerFirst" : "footerFirst" };
    if (even && evenB) return { blocks: evenB, source: kind === "header" ? "headerEven" : "footerEven" };
    const base = ps[kind];
    return base ? { blocks: base, source: kind } : null;
  };

  if (pageSections.some((s) => BAND_KEYS.some((k) => s[k]))) {
    const bandCache = new PrepareCache();
    for (const pg of pages) {
      const ps = pageSections[pg.index]!;
      const cw = ps.pageWidthPx - ps.marginPx.left - ps.marginPx.right;
      const cx = ps.marginPx.left;
      const firstOfSection = pg.index === 0 || pageSections[pg.index - 1] !== ps;
      const pageNum = pageNumbers[pg.index]!;
      const even = pageNum % 2 === 0;
      const header = pickBand(ps, "header", firstOfSection, even);
      if (header) {
        const raw = rawBand === "header";
        const band = layoutBandCached(header.blocks, cw, cx, pageNum, pages.length, bandCache, raw);
        const y0 = headerBandTop(ps, band.height);
        pg.header = band.placed.map((b) => shiftPlaced(b, y0));
        pg.headerSource = header.source;
      }
      const footer = pickBand(ps, "footer", firstOfSection, even);
      if (footer) {
        const raw = rawBand === "footer";
        const band = layoutBandCached(footer.blocks, cw, cx, pageNum, pages.length, bandCache, raw);
        const y0 = footerBandTop(ps, band.height);
        pg.footer = band.placed.map((b) => shiftPlaced(b, y0));
        pg.footerSource = footer.source;
      }
    }
  }

  return {
    pages,
    pageWidthPx: doc.section.pageWidthPx,
    pageHeightPx: doc.section.pageHeightPx,
    marginPx: doc.section.marginPx,
  };
}

// ---------------------------------------------------------------------------
// Margin-band layout: simple vertical stacking (no pagination — a band that
// outgrows its margin overlaps; Word would push content instead, TODO).

function formatPageNumber(n: number, fmt: string | undefined): string {
  switch (fmt) {
    case "roman": return formatListNumber(n, "lowerRoman");
    case "Roman": return formatListNumber(n, "upperRoman");
    case "alpha": return formatListNumber(n, "lowerLetter");
    case "Alpha": return formatListNumber(n, "upperLetter");
    default: return String(n);
  }
}

function substituteTokens(s: string, page: number, pages: number): string {
  return s
    .replace(/\{page(?::(roman|Roman|alpha|Alpha))?\}/g, (_, fmt: string | undefined) => formatPageNumber(page, fmt))
    .replace(/\{pages\}/g, String(pages))
    .replace(/\{date\}/g, () => new Date().toLocaleDateString())
    .replace(/\{time\}/g, () => new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }));
}

/** All run texts in a band's blocks (paragraphs + cell paragraphs at any table
 *  nesting depth) — used to detect volatile {date}/{time} tokens that must bypass
 *  the band cache. */
function blockTexts(blocks: Block[]): string[] {
  const out: string[] = [];
  const visit = (bs: Block[]): void => {
    for (const b of bs) {
      if (b.kind === "paragraph") for (const r of b.runs) out.push(r.text);
      else if (b.kind === "table") for (const row of b.rows) for (const cell of row.cells) visit(cell.blocks);
    }
  };
  visit(blocks);
  return out;
}

function substituteBlock(b: Block, page: number, pages: number): Block {
  const subPara = (p: Paragraph): Paragraph => {
    if (!p.runs.some((r) => r.text.includes("{"))) return p;
    return {
      ...p,
      id: `${p.id}@p${page}`, // per-page identity keeps the band prepare-cache honest
      runs: p.runs.map((r) => ({ ...r, text: substituteTokens(r.text, page, pages) })),
    };
  };
  if (b.kind === "paragraph") return subPara(b);
  if (b.kind === "table") {
    return {
      ...b,
      rows: b.rows.map((row) => ({
        cells: row.cells.map((cell) => ({
          ...cell,
          blocks: cell.blocks.map((cb) => (cb.kind === "paragraph" ? subPara(cb) : cb)),
        })),
      })),
    };
  }
  return b;
}

function layoutBand(
  blocks: Block[],
  width: number,
  originX: number,
  pageNum: number,
  pageCount: number,
  bandCache: PrepareCache,
  rawMode: boolean,
): { placed: PlacedBlock[]; height: number } {
  const getBandLines = (p: Paragraph, w: number): LineBox[] => paragraphLines(p, w, bandCache);
  const placed: PlacedBlock[] = [];
  let y = 0;
  for (const original of blocks) {
    const b = rawMode ? original : substituteBlock(original, pageNum, pageCount);
    if (b.kind === "paragraph") {
      y += b.style.spaceBeforePx;
      const lines = getBandLines(b, width);
      const bandPara: PlacedBlock = { blockId: b.id, x: originX + b.style.indentLeftPx, y, firstLineIndex: 0, lines };
      const bandDecor = paraDecorFor(b.style, width - b.style.indentLeftPx - (b.style.indentRightPx ?? 0), totalLinesHeight(lines));
      if (bandDecor) bandPara.paraDecor = bandDecor;
      placed.push(bandPara);
      y += totalLinesHeight(lines) + b.style.spaceAfterPx;
    } else if (b.kind === "image") {
      const slack = width - b.widthPx;
      const x = originX + (b.align === "center" ? slack / 2 : b.align === "right" ? slack : 0);
      placed.push({
        blockId: b.id,
        x,
        y,
        firstLineIndex: 0,
        lines: [],
        image: { src: b.src, width: b.widthPx, height: b.heightPx, ...(b.crop ? { crop: b.crop } : {}) },
      });
      y += b.heightPx;
    } else if (b.kind === "equation") {
      const box = eqBox(b);
      const align = b.align ?? "center";
      const slack = width - box.width;
      const x = originX + (align === "center" ? slack / 2 : align === "right" ? slack : 0);
      placed.push({ blockId: b.id, x, y, firstLineIndex: 0, lines: [], equation: { box, width: box.width, height: box.ascent + box.descent, baseline: box.ascent } });
      y += box.ascent + box.descent;
    } else {
      const m = measureTable(b, width, getBandLines, EMPTY_LIST_CTX);
      placed.push(placeTable(b, m.rows, m.colWidths, originX + m.xOffset, y, m.tableWidth, 0, EMPTY_LIST_CTX));
      y += m.height;
    }
  }
  return { placed, height: y };
}

function shiftPlaced(p: PlacedBlock, dy: number): PlacedBlock {
  const out: PlacedBlock = { ...p, y: p.y + dy };
  if (p.table) {
    out.table = {
      ...p.table,
      y: p.table.y + dy,
      rows: p.table.rows.map((row) => ({
        ...row,
        y: row.y + dy,
        cells: row.cells.map((cell) => ({
          ...cell,
          y: cell.y + dy,
          blocks: cell.blocks.map((cb) => shiftPlaced(cb, dy)),
        })),
      })),
    };
  }
  return out;
}

// ---------------------------------------------------------------------------
// Tables: equal column widths, padded cells of stacked paragraphs, atomic on
// the page. TODO: row-level page breaking for long tables.

const ATOMIC_GAP = 12; // vertical breathing room around images/tables

// Cell inner padding. Word's default cell margins (TableNormal) are 0 top/bottom
// and 108 twips (~7.2px) left/right — vertical is NOT symmetric with horizontal,
// so a single-line row is only as tall as its line, not line + fixed pad. The
// importer resolves the w:tcMar/w:tblCellMar cascade onto cell.margin; cells
// without one (native/programmatic tables) fall back to these Word defaults.
const DEFAULT_CELL_MARGIN = { top: 0, right: 7.2, bottom: 0, left: 7.2 } as const;
const cellMargin = (c: TableCell): { top: number; right: number; bottom: number; left: number } =>
  c.margin ?? DEFAULT_CELL_MARGIN;

// ---------------------------------------------------------------------------
// Autofit (content-driven) column sizing — only exercised when a table's
// widthMode is "autofitContents"/"autofitWindow". The fixed path (the default)
// never touches any of this.

/** A column's content-width demand: `min` = it cannot be painted narrower than
 *  this without clipping its widest unbreakable token; `max` = its natural width
 *  if nothing wraps. */
interface ColMinMax {
  min: number;
  max: number;
}

/** Width pretext lays the paragraph out at when we want it effectively
 *  unwrapped (one line per hard break) — its natural max content width. */
const AUTOFIT_NATURAL_W = 1_000_000;
/** Floor a solved column at this px so an empty/degenerate column never collapses. */
const AUTOFIT_MIN_COL = 16;

/** Content width of a laid-out line = Σ fragment widths. Alignment-independent
 *  (we sum widths rather than read fragment x), so a centered/right cell measures
 *  the same as a left one. Tab gaps (gapBefore) are not captured — a known minor
 *  gap for tab-heavy autofit cells. */
function lineContentWidth(l: LineBox): number {
  let w = 0;
  for (const f of l.fragments) w += f.width;
  return w;
}

/** Widest laid-out line when the paragraph is broken at `width` (reusing the
 *  cached prepared segments — only the cheap re-break runs). At a huge width this
 *  yields the natural max; at width 1 it yields the widest unbreakable token. */
function paragraphExtent(p: Paragraph, getLines: (p: Paragraph, w: number) => LineBox[], width: number): number {
  let m = 0;
  for (const l of getLines(p, width)) m = Math.max(m, lineContentWidth(l));
  return m;
}

/** Content min/max of one cell BLOCK, in column-width terms (paragraph indents
 *  count toward the required column width; images are treated as unbreakable;
 *  nested tables contribute their own solved-grid min/max). */
function blockMinMax(
  b: Block,
  listCtx: CellListCtx,
  getLines: (p: Paragraph, w: number) => LineBox[],
  available: number,
): ColMinMax {
  if (b.kind === "paragraph") {
    const indent = b.style.indentLeftPx + listCtx.indentOf(b) + (b.style.indentRightPx ?? 0);
    return {
      min: paragraphExtent(b, getLines, 1) + indent,
      max: paragraphExtent(b, getLines, AUTOFIT_NATURAL_W) + indent,
    };
  }
  if (b.kind === "image") {
    // Anchored images are out of flow; in-flow images are unbreakable blocks
    // (the renderer scales them DOWN if a column ends up narrower, so this never
    // forces an overflow — it only expresses the natural-size preference).
    if (b.anchor) return { min: 0, max: 0 };
    return { min: b.widthPx, max: b.widthPx };
  }
  if (b.kind === "equation") {
    const w = eqBox(b).width;
    return { min: w, max: w };
  }
  const cols = autofitColMinMax(b, listCtx, getLines, available);
  return {
    min: cols.reduce((s, c) => s + c.min, 0),
    max: cols.reduce((s, c) => s + c.max, 0),
  };
}

/** Content min/max of a whole cell: the widest block sets each (cell blocks stack
 *  vertically), plus L/R padding, clamped up to any preferred width (w:tcW). */
function cellMinMax(
  cell: TableCell,
  listCtx: CellListCtx,
  getLines: (p: Paragraph, w: number) => LineBox[],
  available: number,
): ColMinMax {
  const mgn = cellMargin(cell);
  const pad = mgn.left + mgn.right;
  const rot = cellRotation(cell.textDirection);
  if (rot) {
    // Vertical-text cell: its column demand is the stack of line heights (the page
    // WIDTH after rotation), independent of the column width — so min === max (the
    // text doesn't wrap into the column). The text LENGTH lands on the row height.
    let cross = 0;
    for (const b of cell.blocks) {
      if (b.kind === "paragraph") {
        cross += b.style.spaceBeforePx + totalLinesHeight(getLines(b, AUTOFIT_NATURAL_W)) + b.style.spaceAfterPx;
      } else if (b.kind === "image") {
        if (!b.anchor) cross += b.heightPx + CELL_BLOCK_GAP;
      } else if (b.kind === "equation") {
        const box = eqBox(b);
        cross += box.ascent + box.descent + CELL_BLOCK_GAP;
      } else {
        cross += measureTable(b, available, getLines, listCtx).height + CELL_BLOCK_GAP;
      }
    }
    let w = cross + pad;
    const prefR = cell.preferredWidth;
    if (prefR) {
      const prefPx = prefR.type === "pct" ? prefR.px * available : prefR.px;
      if (prefPx > w) w = prefPx;
    }
    return { min: w, max: w };
  }
  let min = 0;
  let max = 0;
  for (const b of cell.blocks) {
    const mm = blockMinMax(b, listCtx, getLines, available);
    if (mm.min > min) min = mm.min;
    if (mm.max > max) max = mm.max;
  }
  min += pad;
  max += pad;
  const pref = cell.preferredWidth;
  if (pref) {
    // abs: px is literal. pct: px holds the fraction (0..1) of the table width,
    // resolved against the available width (close enough; the exact table width
    // is circular with the solve).
    const prefPx = pref.type === "pct" ? pref.px * available : pref.px;
    if (prefPx > min) min = prefPx;
    if (prefPx > max) max = prefPx;
  }
  return { min, max };
}

/** Push `demand` into a group of columns, growing each proportionally to its
 *  current width (so wider columns absorb more), but never shrinking. */
function distributeDemand(arr: number[], cols: number[], demand: number): void {
  let cur = 0;
  for (const c of cols) cur += arr[c]!;
  if (demand <= cur) return;
  const extra = demand - cur;
  let wsum = 0;
  for (const c of cols) wsum += arr[c]! + 1;
  for (const c of cols) arr[c]! += extra * ((arr[c]! + 1) / wsum);
}

/** Per-column {min,max} for a table, walking the span-aware grid exactly like
 *  the height pass: span-1 cells set their column directly; spanning cells are
 *  deferred and their demand distributed across the columns they cover. */
function autofitColMinMax(
  t: TableBlock,
  listCtx: CellListCtx,
  getLines: (p: Paragraph, w: number) => LineBox[],
  available: number,
): ColMinMax[] {
  const ncols = effectiveFractions(t).length;
  const colMin = new Array<number>(ncols).fill(0);
  const colMax = new Array<number>(ncols).fill(0);
  const spans: { colStart: number; cols: number[]; mm: ColMinMax }[] = [];
  const rowsRemaining = new Array<number>(ncols).fill(0);
  for (const row of t.rows) {
    let col = 0;
    for (const cell of row.cells) {
      while (col < ncols && rowsRemaining[col]! > 0) col++;
      const span = Math.max(1, cell.colSpan ?? 1);
      const rowSpan = Math.max(1, cell.rowSpan ?? 1);
      const colStart = col;
      const mm = cellMinMax(cell, listCtx, getLines, available);
      if (span === 1) {
        if (colStart < ncols) {
          if (mm.min > colMin[colStart]!) colMin[colStart] = mm.min;
          if (mm.max > colMax[colStart]!) colMax[colStart] = mm.max;
        }
      } else {
        const cols: number[] = [];
        for (let k = 0; k < span && colStart + k < ncols; k++) cols.push(colStart + k);
        if (cols.length > 0) spans.push({ colStart, cols, mm });
      }
      if (rowSpan > 1) for (let k = 0; k < span && colStart + k < ncols; k++) rowsRemaining[colStart + k] = rowSpan;
      col = colStart + span;
    }
    for (let c = 0; c < ncols; c++) if (rowsRemaining[c]! > 0) rowsRemaining[c]!--;
  }
  // Narrower spans first, so a 2-col span settles before a 3-col span layered on it.
  spans.sort((a, b) => a.cols.length - b.cols.length);
  for (const s of spans) {
    distributeDemand(colMin, s.cols, s.mm.min);
    distributeDemand(colMax, s.cols, s.mm.max);
  }
  const out: ColMinMax[] = [];
  for (let c = 0; c < ncols; c++) {
    const min = Math.max(AUTOFIT_MIN_COL, colMin[c]!);
    out.push({ min, max: Math.max(min, colMax[c]!) });
  }
  return out;
}

/** Solve final column widths from per-column min/max and the available width.
 *  CSS automatic-table-layout rules; "autofitWindow" always fills `available`,
 *  "autofitContents" lets the table end up narrower. */
function solveColumns(perCol: ColMinMax[], available: number, mode: "autofitContents" | "autofitWindow"): number[] {
  const ncols = perCol.length;
  if (ncols === 0) return [];
  const min = perCol.map((c) => c.min);
  const max = perCol.map((c) => c.max);
  const sumMin = min.reduce((s, v) => s + v, 0);
  const sumMax = max.reduce((s, v) => s + v, 0);
  if (sumMax <= available) {
    if (mode === "autofitWindow") {
      const slack = available - sumMax;
      const wsum = sumMax > 0 ? sumMax : ncols;
      return max.map((m) => m + slack * ((sumMax > 0 ? m : 1) / wsum));
    }
    return max.slice();
  }
  if (sumMin < available) {
    const slack = available - sumMin;
    const span = sumMax - sumMin;
    return min.map((mn, i) => mn + (span > 0 ? slack * ((max[i]! - mn) / span) : slack / ncols));
  }
  return min.slice(); // overflow: pin to min, content clips/wraps as today
}

type MeasuredCellItem =
  | { kind: "para"; block: Paragraph; lines: LineBox[] }
  | { kind: "image"; block: ImageBlock; width: number; height: number } // scaled to fit the cell
  | { kind: "equation"; block: EquationBlock; box: MathBox; width: number; height: number }
  | { kind: "table"; block: TableBlock; rows: MeasuredRow[]; colWidths: number[]; height: number; tableWidth: number; xOffset: number };

interface MeasuredCell {
  cell: TableCell;
  items: MeasuredCellItem[];
  height: number;
  /** Rendered width — colSpan cells cover several columns. */
  width: number;
  /** First grid column this cell occupies (accounts for rowspan holes above). */
  colStart: number;
  /** Vertical merge: rows this cell covers (1 = normal). */
  rowSpan: number;
  /** Vertical text (w:textDirection): the cell's content was laid out UNWRAPPED in
   *  a swapped frame and is painted rotated 90°. Absent = horizontal. */
  rot?: CellRot;
}
interface MeasuredRow {
  cells: MeasuredCell[];
  height: number;
}

const CELL_BLOCK_GAP = 4; // vertical gap around non-paragraph blocks in cells

/** Vertical-text rotation a cell's w:textDirection requests, or null for the
 *  default horizontal flow. `tbRl` (top→bottom, columns right→left) and the
 *  East-Asian upright variants `tbRlV`/`tbLrV` rotate the text 90° CLOCKWISE;
 *  `btLr` (bottom→top, columns left→right) rotates 90° COUNTER-CLOCKWISE. The
 *  upright `*V` variants degrade to the same 90° rotation (no per-glyph
 *  uprighting); `lrTb`/`lrTbV` stay horizontal. */
type CellRot = "cw" | "ccw";
function cellRotation(dir: TableCell["textDirection"]): CellRot | null {
  switch (dir) {
    case "tbRl":
    case "tbRlV":
    case "tbLrV":
      return "cw";
    case "btLr":
      return "ccw";
    default:
      return null;
  }
}
/** Canvas/pdf rotation angle (radians) for a rotated cell: +π/2 = 90° CW, −π/2 = CCW. */
const cellRotAngle = (rot: CellRot): number => (rot === "cw" ? Math.PI / 2 : -Math.PI / 2);

/** List geometry threaded into the table layout so cell paragraphs can carry a
 *  list marker, just like body paragraphs. `indentOf` is the list-level indent
 *  (added on top of the cell paragraph's own indent); `markers` is the shared
 *  per-paragraph marker map filled by the numbering pass. */
interface CellListCtx {
  indentOf: (p: Paragraph) => number;
  markers: Map<string, { text: string; style: CharStyle; hangingPx: number }>;
}

/** Lists are body-only, so band (header/footer) tables carry no list geometry. */
const EMPTY_LIST_CTX: CellListCtx = { indentOf: () => 0, markers: new Map() };

function measureTable(
  t: TableBlock,
  contentWidth: number,
  getLines: (p: Paragraph, width: number) => LineBox[],
  listCtx: CellListCtx,
): { rows: MeasuredRow[]; colWidths: number[]; height: number; tableWidth: number; xOffset: number } {
  const mode = t.widthMode ?? "fixed";
  let colWidths: number[];
  if (mode === "fixed") {
    const fracs = effectiveFractions(t);
    // A preferred width shrinks the whole grid (columns keep their proportions);
    // absent, the table spans the full content width as it historically did.
    const target = resolveTableWidth(t.preferredWidth, contentWidth, fracs.length);
    colWidths = fracs.map((f) => f * target);
  } else {
    // Content-driven: measure each cell's min/max, solve the grid, THEN run the
    // existing height pass below at the solved widths (two passes over the cells;
    // the expensive prepared segments are cached so only re-breaks repeat).
    const perCol = autofitColMinMax(t, listCtx, getLines, contentWidth);
    colWidths = solveColumns(perCol, contentWidth, mode);
  }
  const ncols = colWidths.length;

  // Grid walk: rowsRemaining[c] > 0 means column c is still covered by a rowspan
  // started in an earlier row, so this row has a "hole" there and its cells shift
  // right past it (HTML-table column assignment). With no rowSpan anywhere this
  // reduces exactly to the old left-to-right placement.
  const rowsRemaining = new Array<number>(ncols).fill(0);
  const rows: MeasuredRow[] = t.rows.map((row) => {
    let col = 0;
    const cells: MeasuredCell[] = row.cells.map((cell) => {
      while (col < ncols && rowsRemaining[col]! > 0) col++;
      const span = Math.max(1, cell.colSpan ?? 1);
      const rowSpan = Math.max(1, cell.rowSpan ?? 1);
      const colStart = col;
      let width = 0;
      for (let k = 0; k < span; k++) width += colWidths[col + k] ?? colWidths[ncols - 1] ?? 40;
      if (rowSpan > 1) for (let k = 0; k < span && col + k < ncols; k++) rowsRemaining[col + k] = rowSpan;
      col += span;
      const mgn = cellMargin(cell);
      const innerWidth = Math.max(8, width - mgn.left - mgn.right);
      const rot = cellRotation(cell.textDirection);
      if (rot) {
        // Vertical-text cell: lay every paragraph out UNWRAPPED (vertical headers
        // don't wrap — the text length sets the row height). After the 90° rotation
        // the laid-out text LENGTH (longest line) becomes the cell's content HEIGHT,
        // so `flow` drives the row; the stack of line heights becomes the cell's
        // content WIDTH (the column demand — see cellMinMax).
        let flow = 0;
        const rotItems: MeasuredCellItem[] = cell.blocks.map((b) => {
          if (b.kind === "paragraph") {
            const lines = getLines(b, AUTOFIT_NATURAL_W);
            let w = 0;
            for (const l of lines) w = Math.max(w, lineContentWidth(l));
            flow = Math.max(flow, w + b.style.indentLeftPx + listCtx.indentOf(b) + (b.style.indentRightPx ?? 0));
            return { kind: "para", block: b, lines };
          }
          if (b.kind === "image") {
            if (b.anchor) return { kind: "image", block: b, width: b.widthPx, height: b.heightPx };
            flow = Math.max(flow, b.widthPx + CELL_BLOCK_GAP);
            return { kind: "image", block: b, width: b.widthPx, height: b.heightPx };
          }
          if (b.kind === "equation") {
            const box = eqBox(b);
            flow = Math.max(flow, box.width + CELL_BLOCK_GAP);
            return { kind: "equation", block: b, box, width: box.width, height: box.ascent + box.descent };
          }
          const m = measureTable(b, innerWidth, getLines, listCtx);
          flow = Math.max(flow, m.tableWidth + CELL_BLOCK_GAP);
          return { kind: "table", block: b, ...m };
        });
        return { cell, items: rotItems, height: flow + mgn.top + mgn.bottom, width, colStart, rowSpan, rot };
      }
      let h = 0;
      const items: MeasuredCellItem[] = cell.blocks.map((b) => {
        if (b.kind === "paragraph") {
          // Wrap at the cell's inner width minus the paragraph's own left/right
          // indent and the list-level indent — mirroring the body wrap width
          // (box.width - indentOf - rightIndentOf). Placement shifts the text
          // right by indentLeftPx + the list indent, so omitting them here would
          // measure at a wider box than is painted and the text would overflow
          // (now clipped) by exactly that indent.
          const avail = innerWidth - b.style.indentLeftPx - listCtx.indentOf(b) - (b.style.indentRightPx ?? 0);
          const lines = getLines(b, Math.max(8, avail));
          h += b.style.spaceBeforePx + totalLinesHeight(lines) + b.style.spaceAfterPx;
          return { kind: "para", block: b, lines };
        }
        if (b.kind === "image") {
          // Anchored (behind/in-front) images are lifted out of the flow: they do
          // NOT advance the cursor or consume cell height — exactly like a
          // body-level anchored image. Carry the native size; placement positions
          // them by their anchor offsets.
          if (b.anchor) return { kind: "image", block: b, width: b.widthPx, height: b.heightPx };
          // Photo-grid cells: scale wide images down to the cell's inner width.
          const scale = Math.min(1, innerWidth / Math.max(1, b.widthPx));
          const w = b.widthPx * scale;
          const ih = b.heightPx * scale;
          h += ih + CELL_BLOCK_GAP;
          return { kind: "image", block: b, width: w, height: ih };
        }
        if (b.kind === "equation") {
          const box = eqBox(b);
          const hh = box.ascent + box.descent;
          h += hh + CELL_BLOCK_GAP;
          return { kind: "equation", block: b, box, width: box.width, height: hh };
        }
        const m = measureTable(b, innerWidth, getLines, listCtx); // nested table
        h += m.height + CELL_BLOCK_GAP;
        return { kind: "table", block: b, ...m };
      });
      return { cell, items, height: h + mgn.top + mgn.bottom, width, colStart, rowSpan };
    });
    for (let c = 0; c < ncols; c++) if (rowsRemaining[c]! > 0) rowsRemaining[c]!--;
    return { cells, height: 0 };
  });

  // Row heights: single-row cells fix their row; a rowspan cell only forces extra
  // height (added to its last row) when its content exceeds the rows it covers.
  const rowHeight = rows.map((r) => Math.max(0, ...r.cells.filter((c) => c.rowSpan === 1).map((c) => c.height)));
  // Fixed/min row height (w:trHeight): "atLeast" grows the row to the height floor;
  // "exact" pins it (content taller than the box is clipped at paint). Applied
  // before the rowspan-overflow pass so a forced-tall row counts toward a span.
  t.rows.forEach((row, ri) => {
    const h = row.props?.height;
    if (!h) return;
    rowHeight[ri] = h.rule === "exact" ? h.value : Math.max(rowHeight[ri]!, h.value);
  });
  rows.forEach((r, ri) => {
    for (const mc of r.cells) {
      if (mc.rowSpan <= 1) continue;
      const endR = Math.min(rows.length - 1, ri + mc.rowSpan - 1);
      let span = 0;
      for (let rr = ri; rr <= endR; rr++) span += rowHeight[rr]!;
      if (mc.height > span) rowHeight[endR]! += mc.height - span;
    }
  });
  rows.forEach((r, ri) => (r.height = rowHeight[ri]!));

  // Fixed tables render at exactly the box width (fractions sum to 1); autofit
  // tables render at the solved column total, which may be narrower (contents) or
  // exactly the box (window).
  const tableWidth =
    mode === "fixed" && !t.preferredWidth ? contentWidth : colWidths.reduce((s, w) => s + w, 0);
  // Alignment only has room to act when the table is narrower than its band — true
  // for a preferred width OR AutoFit-to-Contents; a full-width table never shifts.
  // A table indent (w:tblInd) shifts the whole table right on TOP of that, like a
  // paragraph's left indent (issue #61).
  const xOffset = tableAlignOffset(t.align, contentWidth, tableWidth) + (t.indentPx ?? 0);
  return { rows, colWidths, height: rowHeight.reduce((s, h) => s + h, 0), tableWidth, xOffset };
}

/** Resolve a table's preferred total width to px, clamped to a sane floor and the
 *  available content width. Absent preference → the full content width. */
function resolveTableWidth(
  pref: TableBlock["preferredWidth"],
  contentWidth: number,
  ncols: number,
): number {
  if (!pref) return contentWidth;
  const raw = pref.type === "pct" ? (pref.value / 100) * contentWidth : pref.value;
  const floor = Math.min(contentWidth, Math.max(1, ncols) * 24); // never below ~24px/col
  return Math.max(floor, Math.min(contentWidth, raw));
}

/** Horizontal shift of a table within its band for the given alignment. */
function tableAlignOffset(
  align: TableBlock["align"],
  contentWidth: number,
  tableWidth: number,
): number {
  const slack = contentWidth - tableWidth;
  if (slack <= 0) return 0;
  return align === "center" ? slack / 2 : align === "right" ? slack : 0;
}

function totalLinesHeight(lines: LineBox[]): number {
  let h = 0;
  for (const l of lines) h += l.height;
  return h;
}

function placeTable(
  t: TableBlock,
  rows: MeasuredRow[],
  colWidths: number[],
  x: number,
  y: number,
  width: number,
  firstRowIndex: number,
  listCtx: CellListCtx,
): PlacedBlock {
  // Cumulative grid-column x offsets, so a cell lands at its colStart regardless
  // of rowspan holes in this row; and cumulative row y within the chunk, so a
  // rowspan cell's height = the sum of the rows it covers (clamped to the chunk
  // if a vertical merge straddles a page break).
  const colX = [0];
  for (const w of colWidths) colX.push(colX[colX.length - 1]! + w);
  const rowY = [y];
  for (const row of rows) rowY.push(rowY[rowY.length - 1]! + row.height);
  // RTL visual column order (w:bidiVisual, issue #61): mirror each cell's left edge
  // about the grid so grid column 0 paints at the right. The logical cell order and
  // content are unchanged — only the x placement flips.
  const gridWidth = colX[colX.length - 1] ?? 0;

  const placedRows = [];
  for (let lr = 0; lr < rows.length; lr++) {
    const row = rows[lr]!;
    const ry = rowY[lr]!;
    const cells = [];
    for (const mc of row.cells) {
      const cxStart = colX[mc.colStart] ?? colX[colX.length - 1]!;
      const cx = x + (t.bidiVisual ? gridWidth - cxStart - mc.width : cxStart);
      const endLr = Math.min(rows.length - 1, lr + mc.rowSpan - 1);
      const cellHeight = rowY[endLr + 1]! - ry;
      const blocks: PlacedBlock[] = [];
      const mgn = cellMargin(mc.cell);
      const innerWidth = mc.width - mgn.left - mgn.right;
      if (mc.rot) {
        // Vertical-text cell: place the blocks in a LOCAL (pre-rotation) frame whose
        // local x-axis is the text-advance direction and local y-axis is the stacked
        // lines. paint/geometry map local→page via translate(origin)+rotate(angle):
        //   tbRl (CW)  anchors the TOP-RIGHT inner corner, stacking lines leftward;
        //   btLr (CCW) anchors the BOTTOM-LEFT corner, stacking lines rightward.
        const innerH = cellHeight - mgn.top - mgn.bottom;
        const angle = cellRotAngle(mc.rot);
        const originX = mc.rot === "cw" ? cx + mgn.left + innerWidth : cx + mgn.left;
        const originY = mc.rot === "cw" ? ry + mgn.top : ry + mgn.top + innerH;
        let ly = 0; // local stack axis (→ page width after rotation)
        for (const it of mc.items) {
          if (it.kind === "para") {
            ly += it.block.style.spaceBeforePx;
            const lx = it.block.style.indentLeftPx + listCtx.indentOf(it.block);
            const placedPara: PlacedBlock = { blockId: it.block.id, x: lx, y: ly, firstLineIndex: 0, lines: it.lines };
            // Carry list markers and paragraph shading/border the same as the
            // horizontal path — they're painted under the cell's rotation transform,
            // so their local-frame coords map onto the page correctly.
            const marker = listCtx.markers.get(it.block.id);
            if (marker) placedPara.marker = { text: marker.text, style: marker.style, x: Math.max(0, lx - marker.hangingPx) };
            const cellDecor = paraDecorFor(
              it.block.style,
              innerH - it.block.style.indentLeftPx - listCtx.indentOf(it.block) - (it.block.style.indentRightPx ?? 0),
              totalLinesHeight(it.lines),
            );
            if (cellDecor) placedPara.paraDecor = cellDecor;
            blocks.push(placedPara);
            ly += totalLinesHeight(it.lines) + it.block.style.spaceAfterPx;
          } else if (it.kind === "image") {
            const anchor = it.block.anchor;
            if (anchor) {
              // Anchored (behind/in-front) image: positioned from the cell's content
              // origin by its offsets, carrying behind/front/z — it does NOT advance
              // the stack cursor, mirroring the horizontal cell path.
              const placedImage: PlacedImage = { src: it.block.src, width: it.block.widthPx, height: it.block.heightPx, z: anchor.z ?? 0 };
              if (anchor.behind) placedImage.behind = true;
              else placedImage.front = true;
              blocks.push({ blockId: it.block.id, x: anchor.offsetXPx, y: anchor.offsetYPx, firstLineIndex: 0, lines: [], image: placedImage });
              continue;
            }
            blocks.push({ blockId: it.block.id, x: 0, y: ly, firstLineIndex: 0, lines: [], image: { src: it.block.src, width: it.width, height: it.height } });
            ly += it.height + CELL_BLOCK_GAP;
          } else if (it.kind === "equation") {
            blocks.push({ blockId: it.block.id, x: 0, y: ly, firstLineIndex: 0, lines: [], equation: { box: it.box, width: it.width, height: it.height, baseline: it.box.ascent } });
            ly += it.height + CELL_BLOCK_GAP;
          } else {
            blocks.push(placeTable(it.block, it.rows, it.colWidths, it.xOffset, ly, it.tableWidth, 0, listCtx));
            ly += it.height + CELL_BLOCK_GAP;
          }
        }
        cells.push({
          x: cx,
          y: ry,
          width: mc.width,
          height: cellHeight,
          originRow: firstRowIndex + lr,
          originCol: mc.colStart,
          blocks,
          ...(mc.cell.shading !== undefined ? { shading: mc.cell.shading } : {}),
          ...(mc.cell.borders !== undefined ? { borders: mc.cell.borders } : {}),
          contentClip: { x: cx + mgn.left, y: ry, width: innerWidth, height: cellHeight },
          rotation: { angle, originX, originY },
        });
        continue;
      }
      // Vertical alignment: by default content hugs the top. For center/bottom,
      // offset the whole block stack by the slack between the content height
      // (mc.height already folds in the cell margins) and the cell's painted box,
      // which is taller when the row is forced up by a sibling or a rowSpan.
      const vAlign = mc.cell.vAlign;
      const vSlack = vAlign === "center" || vAlign === "bottom" ? Math.max(0, cellHeight - mc.height) : 0;
      const vOffset = vAlign === "center" ? vSlack / 2 : vAlign === "bottom" ? vSlack : 0;
      let py = ry + mgn.top + vOffset;
      for (const it of mc.items) {
        if (it.kind === "para") {
          py += it.block.style.spaceBeforePx;
          // List paragraphs shift right by the list indent; the marker hangs in
          // that indent, clamped so it stays inside the cell's left padding.
          const lx = cx + mgn.left + it.block.style.indentLeftPx + listCtx.indentOf(it.block);
          const placedPara: PlacedBlock = {
            blockId: it.block.id,
            x: lx,
            y: py,
            firstLineIndex: 0,
            lines: it.lines, // line.y is already block-relative; coords stay consistent
          };
          const marker = listCtx.markers.get(it.block.id);
          if (marker) {
            placedPara.marker = { text: marker.text, style: marker.style, x: Math.max(cx + 1, lx - marker.hangingPx) };
          }
          // Paragraph border/shading box, sized to the cell's content width minus
          // this paragraph's indents (same as body flow, scoped to the cell).
          const cellDecor = paraDecorFor(
            it.block.style,
            innerWidth - it.block.style.indentLeftPx - listCtx.indentOf(it.block) - (it.block.style.indentRightPx ?? 0),
            totalLinesHeight(it.lines),
          );
          if (cellDecor) placedPara.paraDecor = cellDecor;
          blocks.push(placedPara);
          py += totalLinesHeight(it.lines) + it.block.style.spaceAfterPx;
        } else if (it.kind === "image") {
          const anchor = it.block.anchor;
          if (anchor) {
            // Anchored (behind/in-front) cell image: positioned absolutely from the
            // cell's content origin by its offsets, carrying behind/front/z so the
            // z-order pass can lift it into the right layer. Does NOT advance py —
            // text flows over/under it, mirroring placeAnchoredImage at body level.
            const placedImage: PlacedImage = {
              src: it.block.src,
              width: it.block.widthPx,
              height: it.block.heightPx,
              z: anchor.z ?? 0,
            };
            if (anchor.behind) placedImage.behind = true;
            else placedImage.front = true;
            blocks.push({
              blockId: it.block.id,
              x: cx + mgn.left + anchor.offsetXPx,
              y: ry + mgn.top + anchor.offsetYPx,
              firstLineIndex: 0,
              lines: [],
              image: placedImage,
            });
            continue;
          }
          const innerH = cellHeight - mgn.top - mgn.bottom;
          // A lone photo in a cell taller than it fits (e.g. a rowSpan comp-photo
          // column) is filled object-fit:cover — scaled to cover the cell box,
          // centered, and clipped — instead of sitting at the top over blank space.
          if (mc.items.length === 1 && innerH > it.height + 8) {
            const scale = Math.max(innerWidth / it.block.widthPx, innerH / it.block.heightPx);
            const w = it.block.widthPx * scale;
            const h = it.block.heightPx * scale;
            const ix0 = cx + mgn.left;
            const iy0 = ry + mgn.top;
            blocks.push({
              blockId: it.block.id,
              x: ix0 + (innerWidth - w) / 2,
              y: iy0 + (innerH - h) / 2,
              firstLineIndex: 0,
              lines: [],
              image: { src: it.block.src, width: w, height: h, clip: { x: ix0, y: iy0, width: innerWidth, height: innerH } },
            });
            py += it.height + CELL_BLOCK_GAP;
          } else {
            const slack = innerWidth - it.width;
            const ix =
              cx + mgn.left +
              (it.block.align === "center" ? slack / 2 : it.block.align === "right" ? slack : 0);
            blocks.push({
              blockId: it.block.id,
              x: ix,
              y: py,
              firstLineIndex: 0,
              lines: [],
              image: { src: it.block.src, width: it.width, height: it.height },
            });
            py += it.height + CELL_BLOCK_GAP;
          }
        } else if (it.kind === "equation") {
          const align = it.block.align ?? "center";
          const slack = innerWidth - it.width;
          const ix = cx + mgn.left + (align === "center" ? slack / 2 : align === "right" ? slack : 0);
          blocks.push({
            blockId: it.block.id,
            x: ix,
            y: py,
            firstLineIndex: 0,
            lines: [],
            equation: { box: it.box, width: it.width, height: it.height, baseline: it.box.ascent },
          });
          py += it.height + CELL_BLOCK_GAP;
        } else {
          // nested table — placed recursively, read-only inner cells
          blocks.push(placeTable(it.block, it.rows, it.colWidths, cx + mgn.left + it.xOffset, py, it.tableWidth, 0, listCtx));
          py += it.height + CELL_BLOCK_GAP;
        }
      }
      cells.push({
        x: cx,
        y: ry,
        width: mc.width,
        height: cellHeight,
        originRow: firstRowIndex + lr,
        originCol: mc.colStart,
        blocks,
        ...(mc.cell.shading !== undefined ? { shading: mc.cell.shading } : {}),
        ...(mc.cell.borders !== undefined ? { borders: mc.cell.borders } : {}),
        // Horizontal content band (full cell height): clips over-wide text to the
        // inner box so it never paints onto the border or into the next column.
        contentClip: { x: cx + mgn.left, y: ry, width: innerWidth, height: cellHeight },
      });
    }
    placedRows.push({ y: ry, height: row.height, cells });
  }
  const height = rows.reduce((s, r) => s + r.height, 0);
  return {
    blockId: t.id,
    x,
    y,
    firstLineIndex: firstRowIndex, // for tables this is the chunk's first ROW index
    lines: [],
    table: { x, y, width, height, rows: placedRows, colWidths },
  };
}
