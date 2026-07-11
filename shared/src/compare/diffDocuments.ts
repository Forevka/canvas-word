// diffDocuments — the 2-way document compare core. Given an ORIGINAL and a
// REVISED document it produces a single MERGED document plus a ReviewLayer of
// suggestions describing every difference, such that:
//
//   bakeReview(mergedDoc, review, "reject") ≡ original   (drop the revisions)
//   bakeReview(mergedDoc, review, "accept") ≡ revised     (take the revisions)
//
// The merged document obeys the review invariant (REVIEW.md §1): it holds ALL of
// the original's text (deletions present, painted struck-through) PLUS all of the
// revised text spliced in (insertions present, painted underlined). The existing
// accept/reject resolvers then walk each change back to the chosen side — the
// "resolve differences" surface reused by the compare UI.
//
// Isomorphic + DOM-free: the same core runs in the browser editor and headless.
//
// Phase 1 scope: top-level BODY blocks. Matched paragraphs get a word/char-level
// text + format diff; unmatched blocks become add/remove structural records;
// non-paragraph matched blocks (image/table/equation with the same key) are kept
// as-is (their internal props are not yet diffed). 3-way merge is Phase 2.

import { deterministicUserId, type UserInfo } from "../change";
import { DEFAULT_CHAR_STYLE } from "../model/defaults";
import type { Block, Document, Paragraph, ParaStyle, Run } from "../model/document";
import { foldRegistries, reconcileSource } from "../model/merge";
import { normalizeRuns, sliceRuns, styleEq, type Op } from "../model/ops";
import { styleOfCharAt, textOfRuns } from "../model/text";
import { freshId } from "../ids";
import type { ReviewLayer, StructuralChange, Suggestion } from "../review/model";
import { blockKey } from "./blockKey";
import { align } from "./lcs";
import { charStyleDelta, paraStyleDelta } from "./styleDelta";
import { tokenize, type Granularity } from "./tokenize";

export interface CompareOptions {
  /** Attribution for every produced suggestion. Defaults to a synthetic
   *  "Revised Document" author (a stable, deterministic id + color). */
  author?: UserInfo;
  /** docId stamped on the ReviewLayer (match the editor's docId when seeding it). */
  docId?: string;
  /** Intra-paragraph diff granularity. "word" (default) keeps the redline
   *  readable; "char" produces a finer diff. */
  granularity?: Granularity;
  /** Wall-clock base (ms) for suggestion `createdAt`s. The core never reads the
   *  clock (it must run identically headless), so callers that want real
   *  timestamps for display pass `Date.now()`; each record gets base + a
   *  monotonic offset, preserving deterministic newest-first ordering. Omit ⇒ a
   *  0-based counter (deterministic, but a meaningless "time ago"). */
  startTime?: number;
}

export interface CompareStats {
  /** Tracked text insertions (revised-only spans). */
  insertions: number;
  /** Tracked text deletions (original-only spans). */
  deletions: number;
  /** Tracked character-format changes on otherwise-equal text. */
  formatChanges: number;
  /** Whole blocks present only in the revised document. */
  blockInserts: number;
  /** Whole blocks present only in the original document. */
  blockDeletes: number;
  /** Paragraph-level style changes (alignment, spacing, …). */
  paraStyleChanges: number;
}

export interface CompareResult {
  mergedDoc: Document;
  review: ReviewLayer;
  stats: CompareStats;
}

/** The default synthetic author for compare-produced suggestions. */
export function revisedAuthor(): UserInfo {
  return { id: deterministicUserId("Revised Document"), firstName: "Revised", lastName: "Document" };
}

const posAt = (blockId: string, offset: number) => ({ blockId, offset });

/** Compute the diff of `original` → `revised`. Pure; neither input is mutated. */
export function diffDocuments(original: Document, revised: Document, opts: CompareOptions = {}): CompareResult {
  const author = opts.author ?? revisedAuthor();
  const granularity = opts.granularity ?? "word";

  // Reconcile the revised document into the original's id/registry space so its
  // block/style/list/etc. ids can't collide or alias with the original's. The
  // returned sourceBlocks (B′) carry fresh ids + reconciled style references.
  const r = reconcileSource(original, revised);
  const bBlocks = r.sourceBlocks;

  const suggestions: Suggestion[] = [];
  const stats: CompareStats = { insertions: 0, deletions: 0, formatChanges: 0, blockInserts: 0, blockDeletes: 0, paraStyleChanges: 0 };
  // Monotonic createdAt so structural records fold newest-first deterministically
  // (no Date.now in the isomorphic core — it must run identically on the backend).
  // An optional wall-clock base lets a UI caller get real timestamps for display
  // without the core reading the clock.
  const base = opts.startTime ?? 0;
  let seq = 0;
  const now = (): number => base + ++seq;

  const push = (s: Suggestion): void => {
    suggestions.push(s);
  };
  const mkSug = (
    kind: Suggestion["kind"],
    blockId: string,
    start: number,
    end: number,
    extra?: Partial<Suggestion>,
  ): Suggestion => ({ id: freshId(), kind, anchor: { start: posAt(blockId, start), end: posAt(blockId, end) }, author, createdAt: now(), ...extra });

  const mergedBlocks: Block[] = [];
  // Deferred whole-paragraph deletion strikes: only emitted if the merged body
  // ends up with ≥2 blocks, so a removeBlock-on-accept always has a neighbor for
  // its child anchors to collapse onto (avoids a delete-on-a-vanished-block).
  const pendingDeleteStrikes: Array<{ blockId: string; len: number; groupId: string }> = [];

  // ---- whole-block structural helpers ----

  const structural = (op: Op, inverse: Op, blockId: string, applied: boolean, groupId: string): Suggestion => {
    const change: StructuralChange = { op, inverse, blockId, applied };
    return mkSug("structural", blockId, 0, 0, { structural: change, groupId });
  };

  const deleteBlock = (block: Block, index: number): void => {
    mergedBlocks.push(block); // stays live; removed only on accept
    const groupId = freshId();
    // Not-yet-applied removeBlock: accept removes it, reject keeps it.
    push(structural({ type: "removeBlock", blockId: block.id }, { type: "insertBlock", index, block, where: "body" }, block.id, false, groupId));
    if (block.kind === "paragraph") {
      const len = textOfRuns(block.runs).length;
      if (len > 0) pendingDeleteStrikes.push({ blockId: block.id, len, groupId });
    }
    stats.blockDeletes++;
  };

  const insertBlock = (block: Block, index: number): void => {
    mergedBlocks.push(block); // spliced in; removed only on reject
    const groupId = freshId();
    // Already-applied insertBlock: accept keeps it, reject removes it.
    push(structural({ type: "insertBlock", index, block, where: "body" }, { type: "removeBlock", blockId: block.id }, block.id, true, groupId));
    if (block.kind === "paragraph") {
      const len = textOfRuns(block.runs).length;
      if (len > 0) push(mkSug("insert", block.id, 0, len, { groupId }));
    }
    stats.blockInserts++;
  };

  // ---- intra-paragraph diff (matched pair) ----

  const diffParagraph = (a: Paragraph, b: Paragraph): Paragraph => {
    const fallback = a.runs[0]?.style ?? DEFAULT_CHAR_STYLE;
    const runs: Run[] = [];
    let off = 0; // running length of the merged text (merged-coordinate cursor)

    const ta = tokenize(textOfRuns(a.runs), granularity);
    const tb = tokenize(textOfRuns(b.runs), granularity);
    const hunks = align(ta, tb, (t) => t.text);

    const aSpan = (h: { ai: number; a: unknown[] }): [number, number] => [ta[h.ai]!.start, ta[h.ai + h.a.length - 1]!.end];
    const bSpan = (h: { bi: number; b: unknown[] }): [number, number] => [tb[h.bi]!.start, tb[h.bi + h.b.length - 1]!.end];

    const emitDelete = (aStart: number, aEnd: number): void => {
      runs.push(...sliceRuns(a.runs, aStart, aEnd));
      const len = aEnd - aStart;
      push(mkSug("delete", a.id, off, off + len));
      stats.deletions++;
      off += len;
    };
    const emitInsert = (bStart: number, bEnd: number): void => {
      runs.push(...sliceRuns(b.runs, bStart, bEnd));
      const len = bEnd - bStart;
      push(mkSug("insert", a.id, off, off + len));
      stats.insertions++;
      off += len;
    };

    for (const h of hunks) {
      if (h.tag === "equal") {
        // Common text — but the per-character STYLE may differ (a format change).
        const [aStart] = aSpan(h);
        const [bStart] = bSpan(h);
        const common = ta.slice(h.ai, h.ai + h.a.length).map((t) => t.text).join("");
        let g = 0;
        while (g < common.length) {
          const sa = styleOfCharAt(a.runs, aStart + g) ?? fallback;
          const sb = styleOfCharAt(b.runs, bStart + g) ?? fallback;
          const same = styleEq(sa, sb);
          let g2 = g + 1;
          while (g2 < common.length) {
            const na = styleOfCharAt(a.runs, aStart + g2) ?? fallback;
            const nb = styleOfCharAt(b.runs, bStart + g2) ?? fallback;
            if (styleEq(na, nb) !== same) break;
            if (!styleEq(same ? na : nb, same ? sa : sb)) break; // keep the emitted run homogeneous
            g2++;
          }
          const segText = common.slice(g, g2);
          if (same) {
            runs.push({ text: segText, style: sa });
          } else {
            runs.push({ text: segText, style: sb });
            const { patch, inverse } = charStyleDelta(sa, sb);
            push(mkSug("format", a.id, off, off + segText.length, { patch, inverse }));
            stats.formatChanges++;
          }
          off += segText.length;
          g = g2;
        }
      } else if (h.tag === "delete") {
        const [aStart, aEnd] = aSpan(h);
        emitDelete(aStart, aEnd);
      } else if (h.tag === "insert") {
        const [bStart, bEnd] = bSpan(h);
        emitInsert(bStart, bEnd);
      } else {
        // replace: original text struck, revised text inserted (Word order).
        const [aStart, aEnd] = aSpan(h);
        emitDelete(aStart, aEnd);
        const [bStart, bEnd] = bSpan(h);
        emitInsert(bStart, bEnd);
      }
    }

    // Paragraph-level style change (same paragraph, different ParaStyle): an
    // already-applied setParaStyle structural — accept keeps the revised style,
    // reject restores the original.
    const pd = paraStyleDelta(a.style, b.style);
    let style: ParaStyle = a.style;
    if (Object.keys(pd.patch).length > 0) {
      style = { ...a.style, ...pd.patch };
      push(
        structural(
          { type: "setParaStyle", blockId: a.id, patch: pd.patch },
          { type: "setParaStyle", blockId: a.id, patch: pd.inverse },
          a.id,
          true,
          freshId(),
        ),
      );
      stats.paraStyleChanges++;
    }

    const merged: Paragraph = { kind: "paragraph", id: a.id, revision: 0, runs: normalizeRuns(runs, fallback), style };
    if (a.fieldId !== undefined) merged.fieldId = a.fieldId;
    if (a.sdtPath !== undefined) merged.sdtPath = a.sdtPath;
    return merged;
  };

  // ---- block-level walk ----

  // Pair one original block with one revised block: paragraphs get an
  // intra-block text/format diff; anything else stays a replace (delete + insert).
  const pairBlock = (ab: Block, bb: Block): void => {
    if (ab.kind === "paragraph" && bb.kind === "paragraph") mergedBlocks.push(diffParagraph(ab, bb));
    else {
      deleteBlock(ab, mergedBlocks.length);
      insertBlock(bb, mergedBlocks.length);
    }
  };

  const hunks = align(original.blocks, bBlocks, blockKey);
  for (const h of hunks) {
    if (h.tag === "equal") {
      for (let k = 0; k < h.a.length; k++) pairBlock(h.a[k]!, h.b[k]!);
    } else if (h.tag === "delete") {
      for (const ab of h.a) deleteBlock(ab, mergedBlocks.length);
    } else if (h.tag === "insert") {
      for (const bb of h.b) insertBlock(bb, mergedBlocks.length);
    } else {
      // replace: pair original↔revised blocks positionally (so a modified
      // paragraph gets a word-level diff, not a whole-paragraph swap), then any
      // surplus on either side becomes a pure delete / insert.
      const pairs = Math.min(h.a.length, h.b.length);
      for (let k = 0; k < pairs; k++) pairBlock(h.a[k]!, h.b[k]!);
      for (let k = pairs; k < h.a.length; k++) deleteBlock(h.a[k]!, mergedBlocks.length);
      for (let k = pairs; k < h.b.length; k++) insertBlock(h.b[k]!, mergedBlocks.length);
    }
  }

  // Emit the deferred whole-paragraph deletion strikes now the merged length is
  // known (guard: a removeBlock-on-accept needs a surviving neighbor).
  if (mergedBlocks.length >= 2) {
    for (const d of pendingDeleteStrikes) {
      push(mkSug("delete", d.blockId, 0, d.len, { groupId: d.groupId }));
      stats.deletions++;
    }
  }

  const mergedDoc = foldRegistries(original, r, original.section, mergedBlocks);
  const review: ReviewLayer = { docId: opts.docId ?? "compare", baseVersion: 0, suggestions, threads: [] };
  return { mergedDoc, review, stats };
}
