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
import type { Block, CharStyle, Document, ImageBlock, Paragraph, Run, SectionPatch, SectionProps, TableBlock, TableCell, TabStop } from "@cw/shared";
import { effectiveFractions, isHiddenParagraph } from "@cw/shared";
import { formatListNumber, markerText, type ListDefinition, type ListLevel } from "@cw/shared";
import type { InlineFragment, LayoutTree, LineBox, Page, PlacedBlock, PlacedImage } from "./layoutTree";
import { PrepareCache, prepareRuns, type PreparedSegment } from "./prepareCache";
import { charStyleToFont, fontMetrics, measureTextWidth } from "./metrics";

/** Word's default tab interval when a `\t` runs past the last explicit stop
 *  (0.5 inch at 96 dpi). */
const DEFAULT_TAB_STOP_PX = 48;

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

export function createLayoutEngine(): LayoutEngine {
  const prepCache = new PrepareCache();
  // Second cache tier: LineBox[] per (block revision, width). A keystroke
  // re-breaks ONE paragraph; every other block's lines are reused as-is and
  // re-layout degenerates to the pagination walk (pure arithmetic).
  const linesCache = new Map<string, { revision: number; width: number; lines: LineBox[] }>();

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

  return {
    layout(doc: Document, _dirtyBlockIds?: string[], options?: LayoutOptions): LayoutTree {
      const rawBand = options?.rawBand ?? null;
      touched = rawBand === null ? new Set<string>() : null;
      try {
        const tree = layoutDocument(doc, getLines, (p) => {
          touched?.add(p.id);
          return prepCache.get(p);
        }, rawBand);
        if (touched) {
          for (const id of linesCache.keys()) if (!touched.has(id)) linesCache.delete(id);
          prepCache.retainOnly(touched);
        }
        return tree;
      } finally {
        touched = null;
      }
    },
    evict(blockId: string): void {
      prepCache.evict(blockId);
      linesCache.delete(blockId);
    },
    reset(): void {
      prepCache.clear();
      linesCache.clear();
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

/** Default line metrics for a fragment-less line (empty paragraph / segment). */
function emptyLineMetrics(p: Paragraph): { height: number; ascent: number } {
  const style = p.runs[0]?.style;
  const fontSize = style?.fontSizePx ?? 16;
  const m = style ? fontMetrics(charStyleToFont(style)) : { ascent: fontSize * 0.8, descent: fontSize * 0.2 };
  const natural = m.ascent + m.descent;
  const height = Math.max(natural, p.style.lineHeight * fontSize);
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
    const fm = fontMetrics(charStyleToFont(run.style));
    maxAscent = Math.max(maxAscent, fm.ascent);
    maxDescent = Math.max(maxDescent, fm.descent);
    maxFontSize = Math.max(maxFontSize, run.style.fontSizePx);
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
        width: f.occupiedWidth,
      };
      if (m.map) frag.offsetMap = m.map;
      frags.push({ frag, hadGap: f.gapBefore > 0, spaces: countSpaces(f.text) });
    }
    x += f.occupiedWidth;
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
  const height = Math.max(natural, p.style.lineHeight * maxFontSize);
  const leading = (height - natural) / 2; // half-leading above and below
  return { frags, width: line.width, height, ascent: leading + maxAscent, end: line.end };
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
}

/** Next tab stop strictly past `curX`: the first explicit stop, else the next
 *  default-interval multiple. */
function resolveTabStop(curX: number, stops: TabStop[], contentWidth: number): TabStop {
  for (const s of stops) if (s.posPx > curX + 0.5) return s;
  let pos = (Math.floor(curX / DEFAULT_TAB_STOP_PX) + 1) * DEFAULT_TAB_STOP_PX;
  if (pos <= curX) pos = curX + DEFAULT_TAB_STOP_PX;
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
  let curX = firstIndent;
  let maxH = 0;
  let maxA = 0;
  let hasTab = false;

  const flush = (isLast: boolean): void => {
    if (maxH === 0 && baseStyle) {
      const fm = fontMetrics(charStyleToFont(baseStyle));
      maxH = Math.max(fm.ascent + fm.descent, p.style.lineHeight * baseStyle.fontSizePx);
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
      ...(curLeaders.length > 0 ? { leaders: curLeaders } : {}),
    });
    curFrags = [];
    curLeaders = [];
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
      curX = targetX;
      hasTab = true;
    }
    if (pc.runs.length === 0) return; // empty piece (e.g. text before a leading tab)

    // Lay the piece out with wrapping: first sub-line continues from curX; any
    // overflow opens new full-width lines at the left margin.
    const pseg: PreparedSegment = { prepared: prepareRuns(pc.runs), runs: pc.runs, startOffset: pc.start };
    const runStarts = segRunOffsets(pc.runs, pc.start);
    const runCursors = pc.runs.map(() => 0);
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

function paragraphLines(p: Paragraph, contentWidth: number, cache: PrepareCache): LineBox[] {
  const segments = cache.get(p);
  const lines: LineBox[] = [];

  // Phase 1: break lines per SEGMENT (segments = soft-break "\v" pieces).
  // Fragments are line-LOCAL; justification needs to know each segment's last
  // line (it stays ragged, like a paragraph's final line), so alignment is a
  // second pass.
  const raw: RawLine[] = [];

  let first = true; // first line of the PARAGRAPH (first-line indent)
  for (const seg of segments) {
    // Tab path: a segment with hard tabs lays out as one tab-stopped line (the
    // pretext collapse path would otherwise eat the tabs). Dormant unless run
    // text actually carries "\t".
    if (seg.runs.some((r) => r.text.includes("\t"))) {
      raw.push(...layoutTabbedSegment(p, seg.runs, seg.startOffset, contentWidth, first ? p.style.indentFirstLinePx : 0));
      first = false;
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
  }

  // Phase 2: alignment + justification, then final LineBoxes.
  let y = 0;
  for (const rl of raw) {
    const slack = rl.lineMaxWidth - rl.width;
    let startX = rl.indent;
    if (p.style.align === "center") startX += slack / 2;
    else if (p.style.align === "right") startX += slack;
    else if (p.style.align === "justify" && !rl.lastOfSegment) justifyLine(rl.frags, slack);
    for (const rf of rl.frags) rf.frag.x += startX;
    const box: LineBox = { y, height: rl.height, ascent: rl.ascent, fragments: rl.frags.map((rf) => rf.frag) };
    if (rl.emptyOffset !== undefined) box.emptyOffset = rl.emptyOffset;
    if (rl.leaders && rl.leaders.length > 0) {
      box.leaders = rl.leaders.map((l) => ({ ...l, x1: l.x1 + startX, x2: l.x2 + startX }));
    }
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
}

export function resolveSections(doc: Document): ResolvedSection[] {
  const out: ResolvedSection[] = [];
  for (let i = 0; i < doc.blocks.length; i++) {
    const b = doc.blocks[i]!;
    if (b.kind === "paragraph" && b.style.sectionBreak) {
      out.push({ props: effectiveSection(doc.section, b.style.sectionBreak.props), endBlock: i });
    }
  }
  out.push({ props: doc.section, endBlock: doc.blocks.length - 1 });
  return out;
}

function layoutDocument(
  doc: Document,
  getLines: (p: Paragraph, width: number) => LineBox[],
  getPrepared: (p: Paragraph) => PreparedSegment[],
  rawBand: "header" | "footer" | null,
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
          const visible = !!pb.table || !!pb.image || pb.lines.some((l) => l.fragments.some((f) => f.text.trim().length > 0));
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
  // colWidth is the MEASURE width for paragraphs/tables in the section.
  let colCount = 1;
  let colGap = 0;
  let colWidth = 0;
  const applySection = (props: SectionProps): void => {
    sec = props;
    const box = contentBoxOf(sec);
    contentX = sec.marginPx.left;
    contentTop = box.top;
    contentBottom = box.bottom;
    contentWidth = sec.pageWidthPx - sec.marginPx.left - sec.marginPx.right;
    colCount = Math.max(1, sec.columns?.count ?? 1);
    colGap = sec.columns?.gapPx ?? 24;
    colWidth = colCount > 1 ? (contentWidth - (colCount - 1) * colGap) / colCount : contentWidth;
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
  const indentOf = (p: Paragraph): number =>
    p.style.indentLeftPx + (listLevelOf(p)?.indentLeftPx ?? 0);
  const rightIndentOf = (p: Paragraph): number => p.style.indentRightPx ?? 0;
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
    | { kind: "table"; block: TableBlock; rows: MeasuredRow[]; colWidths: number[]; height: number };

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
    } else {
      const t = measureTable(block, colWidth, getLines, cellListCtx);
      measured.push({ kind: "table", block, ...t });
    }
  }
  secIdx = 0;
  applySection(sections[0]!.props);

  // Which section governs each page — bands and paint read per-page geometry.
  const pageSections: SectionProps[] = [];
  const mkPage = (): Page => {
    pageSections.push(sec);
    return {
      index: pageSections.length - 1,
      number: pageSections.length, // provisional; the post-pass sets the real value
      blocks: [],
      widthPx: sec.pageWidthPx,
      heightPx: sec.pageHeightPx,
      marginPx: sec.marginPx,
      contentTopPx: contentTop,
      contentBottomPx: contentBottom,
    };
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
  const colX = (): number => contentX + colIdx * (colWidth + colGap);
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

  /** Section boundary: start the next section's first page. An EMPTY current
   *  page is re-stamped with the new geometry instead of leaving a blank page. */
  const startSectionPage = (): void => {
    if (page.blocks.length === 0) {
      page.widthPx = sec.pageWidthPx;
      page.heightPx = sec.pageHeightPx;
      page.marginPx = sec.marginPx;
      page.contentTopPx = contentTop;
      page.contentBottomPx = contentBottom;
      pageSections[page.index] = sec;
      colIdx = 0;
      colStartCount = 0;
      y = contentTop;
      floats = [];
    } else {
      hardPage();
    }
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
    if (marker) placed.marker = { text: marker.text, style: marker.style, x: placed.x - marker.hangingPx };
    if (block.style.tocEntry) {
      // The number decorates the LAST chunk's last line; right edge = this column's.
      tocPlacements.set(block.id, { chunk: placed, rightEdge: colX() + colWidth, para: block });
    }
    page.blocks.push(placed);
    for (let k = from; k < from + count; k++) {
      const line = lines[k]!;
      placed.lines.push({ ...line, y: y - placed.y });
      y += line.height;
    }
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
        // Widow: never push a single last line to the next page — give it company.
        if (remaining - take === 1) take -= WIDOW_MIN - 1;
        // Orphan: at the paragraph's start, keep ≥2 lines here or move it whole.
        if (i === 0 && take > 0 && take < ORPHAN_MIN) take = 0;
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
    const placedImage: PlacedImage = { src: img.src, width: img.widthPx, height: img.heightPx, z: a.z ?? 0 };
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
      image: { src: img.src, width: img.widthPx, height: img.heightPx },
    });
    if (floating) {
      // Text flows beside the image: register the float, do NOT advance y.
      floats.push({ left: x, right: x + img.widthPx, bottom: y + img.heightPx, side: img.align === "right" ? "right" : "left" });
    } else {
      y += img.heightPx;
    }
  };

  /** Float-affected paragraphs: re-break per line with the float-shrunk width
   *  at the line's own y (pretext's per-line maxWidth makes this cheap). A line
   *  that doesn't fit the page is ROLLED BACK (runCursors snapshot) and
   *  re-broken at full width on the next page. Justify is applied with a
   *  one-line buffer (a line is justified once the NEXT line proves it isn't the
   *  segment's last — which stays ragged, like a paragraph's final line). */
  const placeParagraphFloating = (p: Paragraph): void => {
    const segments = getPrepared(p);
    let first = true;
    let lineIdx = 0;
    let placed: PlacedBlock | null = null;

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
          const firstX = frags[0]?.frag.x ?? 0;
          placed.marker = { text: marker.text, style: marker.style, x: placed.x + firstX - marker.hangingPx };
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
          newPage(); // floats cleared; the line re-breaks at full width
          placed = null;
          continue;
        }
        const slack = box.width - indent - rightIndentOf(p) - bl.width;
        let startX = box.x0 - colX() + indent;
        if (p.style.align === "center") startX += slack / 2;
        else if (p.style.align === "right") startX += slack;
        for (const rf of bl.frags) rf.frag.x += startX;
        // Justify the PREVIOUS line now that this one exists (so it's non-last);
        // buffer this line for the same decision next iteration.
        if (p.style.align === "justify") {
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
          newPage();
          placed = null;
        }
        pushLine([], m.height, m.ascent, seg.startOffset);
        first = false;
      }
    }
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
    let ri = 0;
    while (ri < m.rows.length) {
      let fit = 0;
      let yy = y;
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
        let chunkH = 0;
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
      const chunk = m.rows.slice(ri, ri + fit);
      page.blocks.push(placeTable(m.block, chunk, m.colWidths, colX(), y, colWidth, ri, cellListCtx));
      y += chunk.reduce((s, r) => s + r.height, 0);
      ri += fit;
      if (ri < m.rows.length) newPage();
    }
  };

  for (let bi = 0; bi < measured.length; bi++) {
    // Crossing into a new section: swap geometry, force its first page.
    if (bi > sections[secIdx]!.endBlock) {
      while (bi > sections[secIdx]!.endBlock) applySection(sections[++secIdx]!.props);
      startSectionPage();
    }
    const m = measured[bi]!;

    // Two directly-adjacent atomics (a title-bar table stacked on its data grid,
    // back-to-back images) sit FLUSH — the gap belongs between an atomic and
    // surrounding text, not between atomics, so don't double it into an empty line.
    // Anchored (behind/in-front) images are out of flow — transparent to the
    // gap logic so they neither take a gap nor suppress one between neighbours.
    const atomicKind = (x: Measured | undefined): boolean =>
      x?.kind === "table" || (x?.kind === "image" && x.block.anchor === undefined);
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
    y += block.style.spaceBeforePx;
    if (y >= bottomY() && colHasContent()) newPage();

    // Float-affected paragraphs re-break per line with float-shrunk widths —
    // bypassing the precomputed full-width lines (and widow/orphan rules).
    if (floatsActiveAt(y)) {
      placeParagraphFloating(block);
      y += block.style.spaceAfterPx;
      continue;
    }

    // Keep-with-next: if this block fits here but the keep-CHAIN hanging off it
    // can't start on this page (each keepWithNext member must fit WHOLE, and
    // the first non-keep successor must keep ≥ its orphan/widow take —
    // mirroring placeParagraph's rules), break before this block instead.
    if (block.style.keepWithNext === true && measured[bi + 1] !== undefined && colHasContent()) {
      let yy = y + totalHeight(lines) + block.style.spaceAfterPx;
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
        yy += m2.block.style.spaceBeforePx;
        const inChain = m2.block.style.keepWithNext === true && measured[k + 1] !== undefined;
        if (inChain || m2.block.style.keepLinesTogether === true) {
          // Chain member (or keep-lines paragraph): must fit WHOLE.
          const h = totalHeight(m2.lines);
          if (yy + h > bottomY()) {
            ok = false;
            break;
          }
          yy += h + m2.block.style.spaceAfterPx;
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
          yy += h + m2.block.style.spaceAfterPx;
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
        if (fit < m2.lines.length) {
          if (m2.lines.length - take === 1) take -= WIDOW_MIN - 1; // widow
          if (take < ORPHAN_MIN) take = 0; // orphan
        }
        if (take <= 0) ok = false;
        break;
      }
      if (!ok) newPage();
    }

    placeParagraph(block, lines);
    y += block.style.spaceAfterPx;
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
        pg.blocks.push(placed);
        ny += totalHeight(np.lines) + np.p.style.spaceAfterPx;
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
  // to close the gap the wider token left. Never stale, no relayout, no painter
  // change. Untagged literal "{page}" body text is left alone (fieldId-gated).
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
    const resolveBlockTokens = (b: PlacedBlock, pageNum: number): void => {
      for (const line of b.lines) resolveLineTokens(line, pageNum);
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
        const band = layoutBand(header.blocks, cw, cx, pageNum, pages.length, bandCache, raw);
        const y0 = headerBandTop(ps, band.height);
        pg.header = band.placed.map((b) => shiftPlaced(b, y0));
        pg.headerSource = header.source;
      }
      const footer = pickBand(ps, "footer", firstOfSection, even);
      if (footer) {
        const raw = rawBand === "footer";
        const band = layoutBand(footer.blocks, cw, cx, pageNum, pages.length, bandCache, raw);
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
      placed.push({ blockId: b.id, x: originX + b.style.indentLeftPx, y, firstLineIndex: 0, lines });
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
        image: { src: b.src, width: b.widthPx, height: b.heightPx },
      });
      y += b.heightPx;
    } else {
      const m = measureTable(b, width, getBandLines, EMPTY_LIST_CTX);
      placed.push(placeTable(b, m.rows, m.colWidths, originX, y, width, 0, EMPTY_LIST_CTX));
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

type MeasuredCellItem =
  | { kind: "para"; block: Paragraph; lines: LineBox[] }
  | { kind: "image"; block: ImageBlock; width: number; height: number } // scaled to fit the cell
  | { kind: "table"; block: TableBlock; rows: MeasuredRow[]; colWidths: number[]; height: number };

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
}
interface MeasuredRow {
  cells: MeasuredCell[];
  height: number;
}

const CELL_BLOCK_GAP = 4; // vertical gap around non-paragraph blocks in cells

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
): { rows: MeasuredRow[]; colWidths: number[]; height: number } {
  const fractions = effectiveFractions(t);
  const colWidths = fractions.map((f) => f * contentWidth);
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

  return { rows, colWidths, height: rowHeight.reduce((s, h) => s + h, 0) };
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

  const placedRows = [];
  for (let lr = 0; lr < rows.length; lr++) {
    const row = rows[lr]!;
    const ry = rowY[lr]!;
    const cells = [];
    for (const mc of row.cells) {
      const cx = x + (colX[mc.colStart] ?? colX[colX.length - 1]!);
      const endLr = Math.min(rows.length - 1, lr + mc.rowSpan - 1);
      const cellHeight = rowY[endLr + 1]! - ry;
      const blocks: PlacedBlock[] = [];
      const mgn = cellMargin(mc.cell);
      const innerWidth = mc.width - mgn.left - mgn.right;
      let py = ry + mgn.top;
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
        } else {
          // nested table — placed recursively, read-only inner cells
          blocks.push(placeTable(it.block, it.rows, it.colWidths, cx + mgn.left, py, innerWidth, 0, listCtx));
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
