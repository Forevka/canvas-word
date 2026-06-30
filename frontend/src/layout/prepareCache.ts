// The pretext prepare-cache: prepareRichInline does the expensive measurement
// once per paragraph; invalidated only when the paragraph's revision changes
// (text/style edits). Width changes do NOT invalidate — re-breaking lines is
// cheap arithmetic.
//
// Soft line breaks: a "\v" in run text is a hard break WITHIN the paragraph
// (Shift+Enter). The paragraph splits into segments at every \v; each segment
// gets its own prepare. The \v itself is one UTF-16 unit in the offset space —
// segment k starts at (end of segment k-1) + 1 — so caret math and backspace
// need zero special-casing.

import {
  prepareRichInline,
  type PreparedRichInline,
  type RichInlineItem,
} from "@chenglou/pretext/rich-inline";
import { clearAnalysisCaches, isCJK, setAnalysisLocale } from "@chenglou/pretext/analysis";
import type { CharStyle, Paragraph, Run } from "@cw/shared";
import { charStyleToFont, measureTextWidth } from "./metrics";
import { firstFamilyToken, MATH_FONT_FAMILY } from "../fonts/clones";
import { equationBox } from "./math/equationLayout";
import { SMALL_CAPS_SCALE } from "../paint/paintStyle";

// ---------------------------------------------------------------------------
// Script-fallback tuning (CJK + Arabic). The fallback families are RE-ASSERTED by
// the engine at the top of each layout() pass (a synchronous, await-free span) from
// its per-instance config — exactly like setActiveFontRegistry. So concurrent
// editors / export jobs each see their own config with no shared-mutable-state race
// and no lock: the values are set and read within the same atomic layout span, never
// across an `await`.

// The CJK fallback family active for the CURRENT layout pass (null = leave CJK runs
// untouched, so the browser's system fallback renders them on screen).
let activeCjkFallback: string | null = null;

/** Set the CJK fallback family for the upcoming layout pass. The engine calls this
 *  synchronously at layout() top; `""`/whitespace is treated as "no fallback". */
export function setActiveCjkFallback(family: string | null | undefined): void {
  activeCjkFallback = family && family.trim().length > 0 ? family.trim() : null;
}

// The Arabic fallback family active for the CURRENT layout pass (null = leave Arabic
// runs untouched, so the browser's system fallback renders them on screen).
let activeArabicFallback: string | null = null;

/** Set the Arabic fallback family for the upcoming layout pass. The engine calls
 *  this synchronously at layout() top; `""`/whitespace is treated as "no fallback". */
export function setActiveArabicFallback(family: string | null | undefined): void {
  activeArabicFallback = family && family.trim().length > 0 ? family.trim() : null;
}

// pretext's analyzer locale is PROCESS-GLOBAL inside the library (it can't be made
// per-call), so the engine applies it synchronously at layout() top too — within the
// same await-free span, so concurrent layouts each assert their own. Idempotent per
// value so the editor's per-keystroke relayouts don't thrash the analyzer cache.
let lastLocale: string | undefined;
let lastLocaleApplied = false;
export function applyCjkLocale(locale: string | undefined): void {
  if (lastLocaleApplied && lastLocale === locale) return;
  setAnalysisLocale(locale);
  clearAnalysisCaches();
  lastLocale = locale;
  lastLocaleApplied = true;
}

const hasCjk = (text: string): boolean => {
  for (const ch of text) if (isCJK(ch)) return true;
  return false;
};

/** True when the code point is in any Arabic Unicode block.
 *  Covers Basic Arabic (0600–06FF), Arabic Supplement (0750–077F),
 *  Arabic Extended-B (0870–089F), Arabic Extended-A (08A0–08FF),
 *  Arabic Presentation Forms-A (FB50–FDFF), and
 *  Arabic Presentation Forms-B (FE70–FEFF). */
function isArabic(ch: string): boolean {
  const cp = ch.codePointAt(0);
  if (cp === undefined) return false;
  return (
    (cp >= 0x0600 && cp <= 0x06ff) ||
    (cp >= 0x0750 && cp <= 0x077f) ||
    (cp >= 0x0870 && cp <= 0x089f) ||
    (cp >= 0x08a0 && cp <= 0x08ff) ||
    (cp >= 0xfb50 && cp <= 0xfdff) ||
    (cp >= 0xfe70 && cp <= 0xfeff)
  );
}

const hasArabic = (text: string): boolean => {
  for (const ch of text) if (isArabic(ch)) return true;
  return false;
};

/** Script tag for a character: "arabic" | "cjk" | "other". Used to split runs at
 *  script boundaries when both a CJK and an Arabic fallback are configured. */
type ScriptTag = "arabic" | "cjk" | "other";

function scriptOf(ch: string, cjkFam: string | null, arabicFam: string | null): ScriptTag {
  if (arabicFam && isArabic(ch)) return "arabic";
  if (cjkFam && isCJK(ch)) return "cjk";
  return "other";
}

/** Split runs at CJK ↔ Arabic ↔ other script boundaries, retargeting each script
 *  piece to its configured fallback family. Offset-transparent: the concatenated text
 *  and its length are unchanged, so every downstream offset/itemIndex mapping still
 *  holds. No-op when no fallback fonts are configured or a run carries no CJK/Arabic. */
export function scriptSplitRuns(runs: Run[]): Run[] {
  const cjkFam = activeCjkFallback;
  const arabicFam = activeArabicFallback;
  if (!cjkFam && !arabicFam) return runs;

  const out: Run[] = [];
  let changed = false;
  for (const run of runs) {
    const runFamily = firstFamilyToken(run.style.fontFamily);
    const needsCjk = cjkFam && runFamily !== cjkFam && hasCjk(run.text);
    const needsArabic = arabicFam && runFamily !== arabicFam && hasArabic(run.text);
    if (run.text.length === 0 || (!needsCjk && !needsArabic)) {
      out.push(run);
      continue;
    }
    changed = true;
    let buf = "";
    let bufScript: ScriptTag | null = null;
    const flush = (): void => {
      if (buf.length === 0) return;
      let fontFamily = run.style.fontFamily;
      if (bufScript === "arabic" && arabicFam) fontFamily = arabicFam;
      else if (bufScript === "cjk" && cjkFam) fontFamily = cjkFam;
      out.push({ text: buf, style: fontFamily === run.style.fontFamily ? run.style : { ...run.style, fontFamily } });
      buf = "";
    };
    for (const ch of run.text) {
      const s = scriptOf(ch, cjkFam, arabicFam);
      if (bufScript !== null && s !== bufScript) flush();
      bufScript = s;
      buf += ch;
    }
    flush();
  }
  return changed ? out : runs;
}

export interface PreparedSegment {
  prepared: PreparedRichInline;
  /** The segment's runs (\v stripped) — itemIndex in pretext fragments maps here. */
  runs: Run[];
  /** Global UTF-16 offset of the segment start within the paragraph text. */
  startOffset: number;
}

interface Entry {
  revision: number;
  segments: PreparedSegment[];
}

/** Split runs at "\v" into per-segment run lists with global start offsets. */
export function segmentRuns(runs: Run[]): { runs: Run[]; startOffset: number }[] {
  const segs: { runs: Run[]; startOffset: number }[] = [];
  let cur: Run[] = [];
  let segStart = 0;
  let pos = 0;
  for (const r of runs) {
    let local = 0;
    for (;;) {
      const idx = r.text.indexOf("\v", local);
      if (idx < 0) break;
      const piece = r.text.slice(local, idx);
      if (piece.length > 0) cur.push({ text: piece, style: r.style });
      pos += idx - local;
      segs.push({ runs: cur, startOffset: segStart });
      pos += 1; // the \v itself occupies one offset
      segStart = pos;
      cur = [];
      local = idx + 1;
    }
    const rest = r.text.slice(local);
    if (rest.length > 0) cur.push({ text: rest, style: r.style });
    pos += rest.length;
  }
  segs.push({ runs: cur, startOffset: segStart });
  return segs;
}

function toItems(runs: Run[]): RichInlineItem[] {
  // Hidden runs stay in the list (so item indices line up with seg.runs and the
  // run-offset accounting is unchanged) but contribute EMPTY text to pretext — no
  // fragment, zero width, never painted. Offsets still span the full text.
  return runs.map((run) => {
    // Inline equation: the run is a single U+FFFC reserving the math box width. We
    // keep the sentinel as the item text (pretext drops empty-text items, so a
    // true zero-width placeholder would erase the fragment entirely). pretext's
    // occupied width for a `break:'never'` item is `naturalWidth + extraWidth`, so
    // we set extraWidth to the SIGNED delta `box.width - sentinelW`: occupied then
    // collapses to exactly `box.width` for narrow AND wide equations (the old
    // `max(0, …)` clamp over-reserved whenever the box was narrower than the U+FFFC
    // glyph — see #16). `naturalWidth === sentinelW` since both measure the sentinel
    // through the same context. `break:'never'` keeps it atomic; breakNextLine then
    // swaps the box's metrics in for the line.
    if (run.style.equation && !run.style.hidden) {
      const font = charStyleToFont(run.style);
      const box = equationBox(run.style.equation, MATH_FONT_FAMILY, run.style.fontSizePx);
      const sentinelW = measureTextWidth(run.text, font);
      return { text: run.text, font, break: "never", extraWidth: box.width - sentinelW };
    }
    const item: RichInlineItem = { text: run.style.hidden ? "" : run.text, font: charStyleToFont(run.style) };
    if (!run.style.hidden && run.style.letterSpacingPx !== undefined) item.letterSpacing = run.style.letterSpacingPx;
    return item;
  });
}

// ---------------------------------------------------------------------------
// Case transforms (w:caps / w:smallCaps). Both render letters UPPERCASED; the
// transform is applied to throwaway DISPLAY runs the layout measures and paints,
// never to the model. It is offset-transparent: the per-code-point uppercase only
// substitutes forms whose UTF-16 length is unchanged, so the concatenated text
// length — and therefore every downstream offset/itemIndex/cluster mapping — is
// preserved (mirrors scriptSplitRuns). The baked runs drop the caps/smallCaps
// flags so a second pass (e.g. tab pieces) is a clean no-op.

const r2 = (n: number): number => Math.round(n * 100) / 100;

// Small caps splits a run by reduced/full size; segment by GRAPHEME CLUSTER (not
// code point) so a combining mark stays attached to its base letter — otherwise a
// decomposed "ä" would split into a reduced "A" plus a full-size mark and the
// accent would mis-shape. Matches geometry.ts's cluster granularity.
const capsSegmenter = new Intl.Segmenter(undefined, { granularity: "grapheme" });

/** Length-preserving uppercase of ONE code point: the uppercased form only when it
 *  occupies the same number of UTF-16 units (keeps offsets 1:1); else the original
 *  (rare expanding cases like ß→SS, ﬀ ligatures stay as authored). */
function upperCp(cp: string): string {
  const up = cp.toUpperCase();
  return up.length === cp.length ? up : cp;
}

/** A lowercase letter small-caps should shrink — uppercasing changes it without
 *  changing its UTF-16 length. */
function isSmallCapped(cp: string): boolean {
  const up = cp.toUpperCase();
  return up.length === cp.length && up !== cp;
}

/** Bake w:caps / w:smallCaps into display runs. caps: uppercase every letter at full
 *  size. smallCaps (takes precedence): uppercase every letter, but split the
 *  originally-lowercase letters into reduced-size sub-runs so the painters draw them
 *  smaller with no small-caps-specific code. No-op when no run carries either flag. */
export function applyCaseTransforms(runs: Run[]): Run[] {
  let any = false;
  for (const r of runs) {
    if ((r.style.caps || r.style.smallCaps) && !r.style.equation) {
      any = true;
      break;
    }
  }
  if (!any) return runs;

  const out: Run[] = [];
  for (const run of runs) {
    const { caps, smallCaps } = run.style;
    if ((!caps && !smallCaps) || run.style.equation || run.text.length === 0) {
      out.push(run);
      continue;
    }
    const baseStyle: CharStyle = { ...run.style, caps: undefined, smallCaps: undefined };
    if (!smallCaps) {
      out.push({ text: [...run.text].map(upperCp).join(""), style: baseStyle });
      continue;
    }
    const reducedStyle: CharStyle = { ...baseStyle, fontSizePx: r2(run.style.fontSizePx * SMALL_CAPS_SCALE) };
    let buf = "";
    let bufReduced: boolean | null = null;
    const flush = (): void => {
      if (buf.length === 0) return;
      out.push({ text: buf, style: bufReduced ? reducedStyle : baseStyle });
      buf = "";
    };
    for (const { segment } of capsSegmenter.segment(run.text)) {
      // The cluster's BASE code point decides reduction (combining marks ride along).
      const base = String.fromCodePoint(segment.codePointAt(0)!);
      const reduced = isSmallCapped(base);
      if (bufReduced !== null && reduced !== bufReduced) flush();
      bufReduced = reduced;
      buf += reduced ? [...segment].map(upperCp).join("") : segment;
    }
    flush();
  }
  return out;
}

/** Prepare an arbitrary run list AND return the (case- + CJK-split) run list it was
 *  prepared from. Callers that map pretext itemIndex/offsets back to runs MUST use
 *  the returned `runs` (not the originals) — splitting changes the item count. */
export function prepareRunSegment(runs: Run[]): { prepared: PreparedRichInline; runs: Run[] } {
  const split = scriptSplitRuns(applyCaseTransforms(runs));
  return { prepared: prepareRichInline(toItems(split)), runs: split };
}

/** Prepare an arbitrary run list (used for tab-stop pieces, which are laid out
 *  outside the per-paragraph segment cache). */
export function prepareRuns(runs: Run[]): PreparedRichInline {
  return prepareRunSegment(runs).prepared;
}

export class PrepareCache {
  private map = new Map<string, Entry>();

  get(p: Paragraph): PreparedSegment[] {
    const hit = this.map.get(p.id);
    if (hit && hit.revision === p.revision) return hit.segments;

    const segments: PreparedSegment[] = segmentRuns(p.runs).map((seg) => {
      // Script-split for CJK fallback BEFORE prepare so item indices, fragment
      // styles, and offsets all derive from the same (split) run list.
      return { ...prepareRunSegment(seg.runs), startOffset: seg.startOffset };
    });
    this.map.set(p.id, { revision: p.revision, segments });
    return segments;
  }

  evict(blockId: string): void {
    this.map.delete(blockId);
  }

  /** Drop every entry — used when the whole document is replaced (the importer
   *  re-mints block ids from i0, so a stale entry at the same id+revision would
   *  otherwise leak the previous document's prepared text into the new one). */
  clear(): void {
    this.map.clear();
  }

  /** Drop every entry whose id is NOT in `keep` (paragraphs no longer in the
   *  document). Called after a full layout to bound the cache to live blocks. */
  retainOnly(keep: ReadonlySet<string>): void {
    for (const id of this.map.keys()) if (!keep.has(id)) this.map.delete(id);
  }
}
