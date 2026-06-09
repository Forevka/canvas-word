// A Change is one entry in the document's history: an atomic, ordered group of
// operations a client made. The full history is a base snapshot + the ordered
// Change log; folding applyOp over the log reconstructs any version (see
// ./replay). This is what lets us store edits efficiently instead of duplicating
// the whole document per version.
//
// Only FORWARD ops are stored — each op's inverse is recomputed by applyOp on
// demand (undo), so the log stays half the size of the editor's undo stack.

import type { Op } from "./model/ops";
import type { DocSelection } from "./model/position";

/** Where a change came from. Mirrors the editor's transaction origins minus
 *  "transient" (IME previews never enter the log). undo/redo are recorded as
 *  literal forward ops so the history is faithful. */
export type ChangeOrigin = "typing" | "command" | "paste" | "undo" | "redo";

/** Caller-supplied identity for a collaborator. The embedder owns auth and
 *  decides who can see what — the editor just attributes/presents whatever it's
 *  handed. `id` flows onto changes and the document creator; first/last name are
 *  shown on carets and in the activity panel. */
export interface UserInfo {
  id: string;
  firstName: string;
  lastName: string;
}

/** Display name for a user (handles missing parts gracefully). */
export function userDisplayName(u: Pick<UserInfo, "firstName" | "lastName">): string {
  return `${u.firstName ?? ""} ${u.lastName ?? ""}`.trim() || "Anonymous";
}

/** A stable id derived from a name (POC identity: same name → same id, so the
 *  backend recognizes a returning person). Small FNV-1a hash → short hex; not
 *  cryptographic — the embedder supplies real ids in production. */
export function deterministicUserId(name: string): string {
  const s = name.trim().toLowerCase().replace(/\s+/g, " ");
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return "u" + (h >>> 0).toString(16).padStart(8, "0");
}

/** A stable, readable color for a user/site id (caret + label tint). */
export function colorForId(id: string): string {
  let h = 0;
  for (let i = 0; i < id.length; i++) h = (h * 31 + id.charCodeAt(i)) >>> 0;
  return `hsl(${h % 360}, 70%, 45%)`;
}

export interface Change {
  /** Globally-unique, client-minted id. Idempotency key: re-submitting the same
   *  change (after a reconnect) must not double-apply. */
  id: string;
  /** Which document this change belongs to. */
  docId: string;
  /** The server version (seq) this change was generated against — what the
   *  client had applied when it made the edit. Drives OT rebasing (Phase 4). */
  baseVersion: number;
  /** Canonical position in the total order, assigned by the server on accept.
   *  Tentative-local (the client's own counter) until acknowledged. */
  seq?: number;
  /** The session that authored the change (see shared/ids). */
  siteId: string;
  /** The caller-supplied user id of the author (attribution). Optional: offline /
   *  no-identity sessions omit it. */
  userId?: string;
  /** What kind of edit produced this change (audit/filtering; undo and redo are
   *  recorded as literal forward ops, so this is how they're recognized). Does
   *  not affect replay. */
  origin: ChangeOrigin;
  /** Author wall-clock time (ms since epoch); for display/audit only, never for
   *  ordering — seq is the source of truth. */
  ts: number;
  /** The forward operations, in application order. */
  ops: Op[];
  /** The author's selection after applying — lets a viewer restore the caret. */
  selectionAfter?: DocSelection | null;
}

/** Order a change set by canonical seq (stable for entries sharing/ lacking a
 *  seq — those keep their relative order). */
export function orderBySeq(changes: readonly Change[]): Change[] {
  return changes
    .map((c, i) => [c, i] as const)
    .sort(([a, ai], [b, bi]) => {
      const sa = a.seq ?? Number.MAX_SAFE_INTEGER;
      const sb = b.seq ?? Number.MAX_SAFE_INTEGER;
      return sa !== sb ? sa - sb : ai - bi;
    })
    .map(([c]) => c);
}

/** Every forward op across a change set, flattened in seq order. */
export function changeOps(changes: readonly Change[]): Op[] {
  return orderBySeq(changes).flatMap((c) => c.ops);
}
