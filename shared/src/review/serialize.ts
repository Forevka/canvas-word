// Review-layer serialization — INDEPENDENT of serializeDocument. The overlay is
// persisted and synced on its own channel (never folded into the document
// snapshot), so the docx pipeline and serializeDocument stay review-blind. The
// layer is pure JSON (UserInfo, Run[], strings, numbers), so a JSON round-trip is
// a correct, dependency-free deep clone — mirroring shared/src/persist/serialize.

import type { ReviewLayer } from "./model";

export const REVIEW_SNAPSHOT_VERSION = 1 as const;

export interface SerializedReview {
  version: typeof REVIEW_SNAPSHOT_VERSION;
  layer: ReviewLayer;
}

const clone = <T>(v: T): T => JSON.parse(JSON.stringify(v)) as T;

export function serializeReview(layer: ReviewLayer): SerializedReview {
  return { version: REVIEW_SNAPSHOT_VERSION, layer: clone(layer) };
}

export function parseReview(snap: SerializedReview): ReviewLayer {
  // Future versions branch on snap.version here.
  return clone(snap.layer);
}
