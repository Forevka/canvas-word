// Anchor rebasing + garbage collection. The review layer's anchors are
// offset-precise positions into the live document, so they must travel with
// every edit exactly like bookmarks do (see frontend rebaseBookmarks). The
// editor runs mapReviewAnchors with each applied core op's position-mapper —
// local edits, undo/redo, AND remote OT edits — then gcReviewLayer drops records
// whose text was removed.

import type { DocPosition } from "../model/position";
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
  const suggestions = layer.suggestions.filter((s) => !isCollapsedAnchor(s.anchor));
  if (suggestions.length === layer.suggestions.length) return layer;
  return { ...layer, suggestions };
}

/** Convenience: rebase through one op then GC (the per-op pairing the runtime
 *  uses next to rebaseBookmarks). */
export function rebaseReview(layer: ReviewLayer, mapPosition: MapPosition): ReviewLayer {
  return gcReviewLayer(mapReviewAnchors(layer, mapPosition));
}
