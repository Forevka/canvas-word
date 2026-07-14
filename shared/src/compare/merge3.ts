// merge3 — 3-way (git-style) document merge. Given a common BASE and two
// derived versions MINE and THEIRS, it produces an ordered list of merge
// segments: regions unchanged in both sides stay; a region changed on only one
// side auto-merges that side; a region changed on BOTH sides in the same way
// auto-merges once; a region changed on both sides differently is a CONFLICT the
// user resolves by picking a side. This is the classic diff3 parse, applied at
// the BLOCK level (paragraphs/tables/images keyed by blockKey) rather than lines.
//
// Reuses the same reconcile + block-key machinery as the 2-way compare so the two
// sides' disjoint id/registry spaces fold into one; `assembleMerge` builds the
// final Document from a set of per-conflict choices. Pure + DOM-free.

import type { Block, Document, SectionProps } from "../model/document";
import { foldRegistries, reconcileSource, type Reconciled } from "../model/merge";
import { freshId } from "../ids";
import { blockKey } from "./blockKey";
import { align } from "./lcs";

/** Which version to take for a conflict (or a whole segment). */
export type MergeChoice = "mine" | "theirs" | "both" | "base";

export type MergeSegmentKind = "stable" | "mine" | "theirs" | "bothSame" | "conflict";

/** One region of the merged document in base order. `stable`/`mine`/`theirs`/
 *  `bothSame` are auto-resolved (their `blocks` are final); `conflict` carries all
 *  three sides and stays unresolved until a choice is made. */
export interface MergeSegment {
  kind: MergeSegmentKind;
  /** Auto-resolved segments: the blocks to emit. Conflicts: the DEFAULT side's
   *  blocks (mine) — use `mine`/`theirs`/`base` for the alternatives. */
  blocks: Block[];
  /** conflict only. */
  id?: string;
  base?: Block[];
  mine?: Block[];
  theirs?: Block[];
}

export interface Merge3Options {
  docId?: string;
}

export interface Merge3Stats {
  /** Regions unchanged in both sides. */
  unchanged: number;
  /** Regions auto-merged (changed on exactly one side, or identically on both). */
  autoMerged: number;
  /** Regions changed differently on both sides — need a choice. */
  conflicts: number;
}

export interface Merge3Result {
  segments: MergeSegment[];
  /** The conflict segments, in document order (each has a stable `id`). */
  conflicts: MergeSegment[];
  stats: Merge3Stats;
  /** Retained so `assembleMerge` can fold the unified registries + section. */
  baseSection: SectionProps;
  reconciled: Reconciled;
  baseDoc: Document;
}

/** Base-index → matched other-index map from the LCS equal hunks (−1 = changed). */
function matchMap(base: readonly Block[], other: readonly Block[]): number[] {
  const m = new Array<number>(base.length).fill(-1);
  for (const h of align(base, other, blockKey)) {
    if (h.tag !== "equal") continue;
    for (let k = 0; k < h.a.length; k++) m[h.ai + k] = h.bi + k;
  }
  return m;
}

const keysEqual = (x: readonly Block[], y: readonly Block[]): boolean =>
  x.length === y.length && x.every((b, i) => blockKey(b) === blockKey(y[i]!));

/** Compute the 3-way merge of `base` → (`mine`, `theirs`). Pure. */
export function merge3(base: Document, mine: Document, theirs: Document, _opts: Merge3Options = {}): Merge3Result {
  // Fold both sides' id/registry spaces into the base's: first mine, then theirs
  // on top (so the unified registry in r2 covers base + mine + theirs).
  const r1 = reconcileSource(base, mine);
  const mineBlocks = r1.sourceBlocks;
  const baseWithMine = foldRegistries(base, r1, base.section, base.blocks);
  const r2 = reconcileSource(baseWithMine, theirs);
  const theirsBlocks = r2.sourceBlocks;

  const O = base.blocks;
  const A = mineBlocks;
  const B = theirsBlocks;
  const matchA = matchMap(O, A);
  const matchB = matchMap(O, B);

  const segments: MergeSegment[] = [];
  const stats: Merge3Stats = { unchanged: 0, autoMerged: 0, conflicts: 0 };

  let oi = 0;
  let ai = 0;
  let bi = 0;
  while (oi < O.length || ai < A.length || bi < B.length) {
    // Stable run: base[oi] is matched to mine[ai] AND theirs[bi] at the cursors.
    if (oi < O.length && matchA[oi] === ai && matchB[oi] === bi) {
      const start = oi;
      while (oi < O.length && matchA[oi] === ai && matchB[oi] === bi) {
        oi++;
        ai++;
        bi++;
      }
      segments.push({ kind: "stable", blocks: O.slice(start, oi) });
      stats.unchanged++;
      continue;
    }
    // Unstable region: advance base to the next index matched in BOTH sides whose
    // matched positions are at/after the current mine+theirs cursors.
    let oEnd = oi;
    while (oEnd < O.length && !(matchA[oEnd] !== -1 && matchB[oEnd] !== -1 && matchA[oEnd]! >= ai && matchB[oEnd]! >= bi)) {
      oEnd++;
    }
    const aEnd = oEnd < O.length ? matchA[oEnd]! : A.length;
    const bEnd = oEnd < O.length ? matchB[oEnd]! : B.length;
    const oSlice = O.slice(oi, oEnd);
    const aSlice = A.slice(ai, aEnd);
    const bSlice = B.slice(bi, bEnd);
    oi = oEnd;
    ai = aEnd;
    bi = bEnd;

    if (keysEqual(aSlice, oSlice)) {
      // Mine unchanged here → theirs' change wins.
      if (bSlice.length > 0 || oSlice.length > 0) stats.autoMerged++;
      segments.push({ kind: "theirs", blocks: bSlice });
    } else if (keysEqual(bSlice, oSlice)) {
      // Theirs unchanged → mine's change wins.
      if (aSlice.length > 0 || oSlice.length > 0) stats.autoMerged++;
      segments.push({ kind: "mine", blocks: aSlice });
    } else if (keysEqual(aSlice, bSlice)) {
      // Both made the same change → take it once.
      stats.autoMerged++;
      segments.push({ kind: "bothSame", blocks: aSlice });
    } else {
      // Diverging changes → a conflict (defaults to mine when assembled).
      stats.conflicts++;
      segments.push({ kind: "conflict", id: freshId(), blocks: aSlice, base: oSlice, mine: aSlice, theirs: bSlice });
    }
  }

  return {
    segments,
    conflicts: segments.filter((s) => s.kind === "conflict"),
    stats,
    baseSection: base.section,
    reconciled: r2,
    baseDoc: base,
  };
}

/** Blocks a conflict contributes for a given choice. */
function conflictBlocks(seg: MergeSegment, choice: MergeChoice): Block[] {
  switch (choice) {
    case "mine":
      return seg.mine ?? [];
    case "theirs":
      return seg.theirs ?? [];
    case "base":
      return seg.base ?? [];
    case "both":
      return [...(seg.mine ?? []), ...(seg.theirs ?? [])];
  }
}

/** Assemble the merged Document from the segments + per-conflict choices
 *  (unspecified conflicts default to "mine"). Folds the unified registries. */
export function assembleMerge(result: Merge3Result, choices: Record<string, MergeChoice> = {}): Document {
  const blocks: Block[] = [];
  for (const seg of result.segments) {
    if (seg.kind === "conflict") blocks.push(...conflictBlocks(seg, choices[seg.id!] ?? "mine"));
    else blocks.push(...seg.blocks);
  }
  return foldRegistries(result.baseDoc, result.reconciled, result.baseSection, blocks);
}
