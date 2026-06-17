// Anchor rebasing + garbage collection. The review layer's anchors are
// offset-precise positions into the live document, so they must travel with
// every edit exactly like bookmarks do (see frontend rebaseBookmarks). The
// editor runs mapReviewAnchors with each applied core op's position-mapper —
// local edits, undo/redo, AND remote OT edits — then gcReviewLayer drops records
// whose text was removed.

import type { DocPosition } from "../model/position";
import type { Document } from "../model/document";
import { blockExists } from "../model/text";
import type { ReviewAnchor, ReviewLayer } from "./model";
import { isCollapsedAnchor } from "./model";

type MapPosition = (p: DocPosition) => DocPosition;

function mapAnchor(a: ReviewAnchor, mapPosition: MapPosition): ReviewAnchor {
  return { start: mapPosition(a.start), end: mapPosition(a.end) };
}

/** Rebase every suggestion + thread anchor through one core op's mapPosition.
 *  Pure; returns a new layer (cheap — only anchors change). */
export function mapReviewAnchors(layer: ReviewLayer, mapPosition: MapPosition): ReviewLayer {
  return {
    ...layer,
    suggestions: layer.suggestions.map((s) => ({ ...s, anchor: mapAnchor(s.anchor, mapPosition) })),
    threads: layer.threads.map((t) => ({ ...t, anchor: mapAnchor(t.anchor, mapPosition) })),
  };
}

/** Fold a sequence of mapPositions over the layer (late-join / batch fast-forward
 *  of anchors through core ops the receiver applied since baseVersion). */
export function mapReviewAnchorsAll(layer: ReviewLayer, mappers: readonly MapPosition[]): ReviewLayer {
  let l = layer;
  for (const m of mappers) l = mapReviewAnchors(l, m);
  return l;
}

/** Drop dead suggestions: a suggestion whose anchor collapsed to a point means
 *  its underlying text was removed by an edit, so the record can never resolve to
 *  anything — remove it. Threads are NOT GC'd by length: a collapsed thread
 *  anchor is a legitimate point comment (and removeBlock maps it to a neighbor,
 *  same as bookmarks), so comments survive. */
export function gcReviewLayer(layer: ReviewLayer): ReviewLayer {
  // structural records carry no text range — their anchor is a degenerate point
  // on a block, so a collapsed anchor is NORMAL for them, not death. They are
  // GC'd by block existence instead (gcStructuralReviewLayer), so skip them here.
  const suggestions = layer.suggestions.filter((s) => s.kind === "structural" || !isCollapsedAnchor(s.anchor));
  if (suggestions.length === layer.suggestions.length) return layer;
  return { ...layer, suggestions };
}

/** Drop structural records whose anchored block no longer exists in the doc
 *  (e.g. a paste's inserted block was rejected/removed, killing the record).
 *  Keyed on whole-block identity, not a text offset, so it needs the live doc —
 *  the runtime runs it after applying core ops, next to gcReviewLayer. */
export function gcStructuralReviewLayer(layer: ReviewLayer, doc: Document): ReviewLayer {
  const suggestions = layer.suggestions.filter(
    (s) => s.kind !== "structural" || !s.structural || blockExists(doc, s.structural.blockId),
  );
  if (suggestions.length === layer.suggestions.length) return layer;
  return { ...layer, suggestions };
}

/** Convenience: rebase through one op then GC (the per-op pairing the runtime
 *  uses next to rebaseBookmarks). */
export function rebaseReview(layer: ReviewLayer, mapPosition: MapPosition): ReviewLayer {
  return gcReviewLayer(mapReviewAnchors(layer, mapPosition));
}
