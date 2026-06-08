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
import type { Paragraph, Run } from "../model/document";
import { charStyleToFont } from "./metrics";

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
  return runs.map((run) => {
    const item: RichInlineItem = { text: run.text, font: charStyleToFont(run.style) };
    if (run.style.letterSpacingPx !== undefined) item.letterSpacing = run.style.letterSpacingPx;
    return item;
  });
}

/** Prepare an arbitrary run list (used for tab-stop pieces, which are laid out
 *  outside the per-paragraph segment cache). */
export function prepareRuns(runs: Run[]): PreparedRichInline {
  return prepareRichInline(toItems(runs));
}

export class PrepareCache {
  private map = new Map<string, Entry>();

  get(p: Paragraph): PreparedSegment[] {
    const hit = this.map.get(p.id);
    if (hit && hit.revision === p.revision) return hit.segments;

    const segments: PreparedSegment[] = segmentRuns(p.runs).map((seg) => ({
      prepared: prepareRichInline(toItems(seg.runs)),
      runs: seg.runs,
      startOffset: seg.startOffset,
    }));
    this.map.set(p.id, { revision: p.revision, segments });
    return segments;
  }

  evict(blockId: string): void {
    this.map.delete(blockId);
  }
}
