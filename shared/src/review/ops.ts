// Review-layer mutation. Unlike core ops, these never move text — they are
// record-level upserts/deletes keyed by id, so any two ReviewOps commute (OT
// between them is the identity) and re-delivery is idempotent. The clean
// separation that lets the review layer be an extension instead of a fork of the
// OT core: CORE ops move text and rebase review anchors (see ./rebase); REVIEW
// ops mutate records and never move text. Neither transforms against the other.
//
// applyReviewOp returns the new layer plus the op's exact INVERSE, mirroring the
// core applyOp contract, so the editor's (separate) review-undo channel is free.

import type { DocPosition } from "../model/position";
import type { UserInfo } from "../change";
import type { Comment, CommentThread, Fragment, ReviewLayer, Suggestion, ThreadStatus } from "./model";

export type ReviewOp =
  | { type: "addSuggestion"; s: Suggestion }
  | { type: "removeSuggestion"; id: string }
  | { type: "growSuggestion"; id: string; end: DocPosition }
  | { type: "addThread"; t: CommentThread }
  | { type: "removeThread"; id: string }
  | { type: "addComment"; threadId: string; c: Comment }
  | { type: "removeComment"; threadId: string; commentId: string }
  | { type: "editComment"; threadId: string; commentId: string; body: Fragment }
  | { type: "setThreadStatus"; threadId: string; status: ThreadStatus };

/** A review op as it travels for sync/persist: tagged with the core document
 *  version its anchor depends on (causal delivery — the receiver holds it until
 *  its doc has reached that version) and the author. Shared by the frontend sync
 *  client and the backend review store. */
export interface ReviewOpEnvelope {
  op: ReviewOp;
  dependsOnSeq: number;
  author?: UserInfo;
}

export interface ReviewApplyResult {
  layer: ReviewLayer;
  /** Exact inverse for the review-undo channel; a no-op inverse (e.g. removing a
   *  record that was already absent) round-trips to the same op so undo is safe. */
  inverse: ReviewOp;
}

const findSuggestion = (l: ReviewLayer, id: string): Suggestion | undefined =>
  l.suggestions.find((s) => s.id === id);
const findThread = (l: ReviewLayer, id: string): CommentThread | undefined =>
  l.threads.find((t) => t.id === id);

/** Upsert a suggestion by id (replace if present → idempotent on re-delivery). */
function putSuggestion(l: ReviewLayer, s: Suggestion): ReviewLayer {
  const i = l.suggestions.findIndex((x) => x.id === s.id);
  const suggestions = l.suggestions.slice();
  if (i >= 0) suggestions[i] = s;
  else suggestions.push(s);
  return { ...l, suggestions };
}

function putThread(l: ReviewLayer, t: CommentThread): ReviewLayer {
  const i = l.threads.findIndex((x) => x.id === t.id);
  const threads = l.threads.slice();
  if (i >= 0) threads[i] = t;
  else threads.push(t);
  return { ...l, threads };
}

function replaceThread(l: ReviewLayer, id: string, fn: (t: CommentThread) => CommentThread): ReviewLayer {
  const i = l.threads.findIndex((x) => x.id === id);
  if (i < 0) return l;
  const threads = l.threads.slice();
  threads[i] = fn(threads[i]!);
  return { ...l, threads };
}

export function applyReviewOp(layer: ReviewLayer, op: ReviewOp): ReviewApplyResult {
  switch (op.type) {
    case "addSuggestion": {
      const prev = findSuggestion(layer, op.s.id);
      return {
        layer: putSuggestion(layer, op.s),
        // Re-adding over an existing record: inverse restores the prior one.
        inverse: prev ? { type: "addSuggestion", s: prev } : { type: "removeSuggestion", id: op.s.id },
      };
    }

    case "removeSuggestion": {
      const prev = findSuggestion(layer, op.id);
      if (!prev) return { layer, inverse: { type: "removeSuggestion", id: op.id } };
      return {
        layer: { ...layer, suggestions: layer.suggestions.filter((s) => s.id !== op.id) },
        inverse: { type: "addSuggestion", s: prev },
      };
    }

    case "growSuggestion": {
      const prev = findSuggestion(layer, op.id);
      if (!prev) return { layer, inverse: { type: "growSuggestion", id: op.id, end: op.end } };
      const oldEnd = prev.anchor.end;
      return {
        layer: putSuggestion(layer, { ...prev, anchor: { ...prev.anchor, end: op.end } }),
        inverse: { type: "growSuggestion", id: op.id, end: oldEnd },
      };
    }

    case "addThread": {
      const prev = findThread(layer, op.t.id);
      return {
        layer: putThread(layer, op.t),
        inverse: prev ? { type: "addThread", t: prev } : { type: "removeThread", id: op.t.id },
      };
    }

    case "removeThread": {
      const prev = findThread(layer, op.id);
      if (!prev) return { layer, inverse: { type: "removeThread", id: op.id } };
      return {
        layer: { ...layer, threads: layer.threads.filter((t) => t.id !== op.id) },
        inverse: { type: "addThread", t: prev },
      };
    }

    case "addComment": {
      const t = findThread(layer, op.threadId);
      if (!t) return { layer, inverse: { type: "removeComment", threadId: op.threadId, commentId: op.c.id } };
      // Idempotent: a re-delivered comment replaces in place, never duplicates.
      const exists = t.comments.some((c) => c.id === op.c.id);
      const comments = exists
        ? t.comments.map((c) => (c.id === op.c.id ? op.c : c))
        : [...t.comments, op.c];
      return {
        layer: replaceThread(layer, op.threadId, (th) => ({ ...th, comments })),
        inverse: { type: "removeComment", threadId: op.threadId, commentId: op.c.id },
      };
    }

    case "removeComment": {
      const t = findThread(layer, op.threadId);
      const prev = t?.comments.find((c) => c.id === op.commentId);
      if (!t || !prev) {
        return { layer, inverse: { type: "removeComment", threadId: op.threadId, commentId: op.commentId } };
      }
      return {
        layer: replaceThread(layer, op.threadId, (th) => ({
          ...th,
          comments: th.comments.filter((c) => c.id !== op.commentId),
        })),
        inverse: { type: "addComment", threadId: op.threadId, c: prev },
      };
    }

    case "editComment": {
      const t = findThread(layer, op.threadId);
      const prev = t?.comments.find((c) => c.id === op.commentId);
      if (!t || !prev) {
        return { layer, inverse: { type: "editComment", threadId: op.threadId, commentId: op.commentId, body: op.body } };
      }
      return {
        layer: replaceThread(layer, op.threadId, (th) => ({
          ...th,
          comments: th.comments.map((c) => (c.id === op.commentId ? { ...c, body: op.body } : c)),
        })),
        inverse: { type: "editComment", threadId: op.threadId, commentId: op.commentId, body: prev.body },
      };
    }

    case "setThreadStatus": {
      const t = findThread(layer, op.threadId);
      if (!t) return { layer, inverse: { type: "setThreadStatus", threadId: op.threadId, status: op.status } };
      return {
        layer: replaceThread(layer, op.threadId, (th) => ({ ...th, status: op.status })),
        inverse: { type: "setThreadStatus", threadId: op.threadId, status: t.status },
      };
    }
  }
}

/** Apply a run of review ops, returning the final layer (inverses discarded). */
export function applyReviewOps(layer: ReviewLayer, ops: readonly ReviewOp[]): ReviewLayer {
  let l = layer;
  for (const op of ops) l = applyReviewOp(l, op).layer;
  return l;
}
