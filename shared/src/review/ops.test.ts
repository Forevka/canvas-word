import { describe, expect, it } from "vitest";
import type { UserInfo } from "../change";
import type { Comment, CommentThread, Suggestion } from "./model";
import { emptyReview } from "./model";
import { applyReviewOp, applyReviewOps, type ReviewOp } from "./ops";

const alice: UserInfo = { id: "u-alice", firstName: "Alice", lastName: "A" };
const bob: UserInfo = { id: "u-bob", firstName: "Bob", lastName: "B" };

const pos = (blockId: string, offset: number) => ({ blockId, offset });
const sug = (id: string, kind: Suggestion["kind"], s = 0, e = 3): Suggestion => ({
  id,
  kind,
  anchor: { start: pos("p", s), end: pos("p", e) },
  author: alice,
  createdAt: 1,
});
const cmt = (id: string, text: string): Comment => ({
  id,
  author: alice,
  body: [{ text, style: { fontFamily: "G", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" } }],
  createdAt: 1,
});
const thread = (id: string, ...comments: Comment[]): CommentThread => ({
  id,
  anchor: { start: pos("p", 0), end: pos("p", 4) },
  status: "open",
  comments,
});

const roundtrip = (start: ReturnType<typeof emptyReview>, op: ReviewOp): void => {
  const { layer, inverse } = applyReviewOp(start, op);
  const back = applyReviewOp(layer, inverse).layer;
  expect(back).toEqual(start);
};

describe("applyReviewOp — apply + inverse round-trip", () => {
  it("addSuggestion / removeSuggestion are inverses", () => {
    const r0 = emptyReview("d");
    const s = sug("s1", "insert");
    const r1 = applyReviewOp(r0, { type: "addSuggestion", s }).layer;
    expect(r1.suggestions).toHaveLength(1);
    roundtrip(r0, { type: "addSuggestion", s });
    roundtrip(r1, { type: "removeSuggestion", id: "s1" });
  });

  it("addSuggestion is idempotent on re-delivery (no duplicates)", () => {
    const s = sug("s1", "delete");
    const r = applyReviewOps(emptyReview("d"), [
      { type: "addSuggestion", s },
      { type: "addSuggestion", s },
    ]);
    expect(r.suggestions).toHaveLength(1);
  });

  it("removeSuggestion of an absent record is a no-op that round-trips", () => {
    const r0 = emptyReview("d");
    const { layer } = applyReviewOp(r0, { type: "removeSuggestion", id: "missing" });
    expect(layer).toEqual(r0);
  });

  it("growSuggestion extends the anchor end and inverts to the old end", () => {
    const r0 = applyReviewOp(emptyReview("d"), { type: "addSuggestion", s: sug("s1", "insert", 0, 3) }).layer;
    const { layer, inverse } = applyReviewOp(r0, { type: "growSuggestion", id: "s1", end: pos("p", 7) });
    expect(layer.suggestions[0]!.anchor.end).toEqual(pos("p", 7));
    expect(applyReviewOp(layer, inverse).layer).toEqual(r0);
  });

  it("thread + comment lifecycle round-trips", () => {
    const r0 = emptyReview("d");
    const t = thread("t1", cmt("c1", "first"));
    const r1 = applyReviewOp(r0, { type: "addThread", t }).layer;
    roundtrip(r0, { type: "addThread", t });
    roundtrip(r1, { type: "addComment", threadId: "t1", c: cmt("c2", "reply") });
    roundtrip(r1, { type: "setThreadStatus", threadId: "t1", status: "resolved" });
    roundtrip(r1, { type: "editComment", threadId: "t1", commentId: "c1", body: cmt("x", "edited").body });
  });

  it("addComment is idempotent (re-delivered reply replaces in place)", () => {
    const r1 = applyReviewOp(emptyReview("d"), { type: "addThread", t: thread("t1", cmt("c1", "first")) }).layer;
    const reply = cmt("c2", "reply");
    const r = applyReviewOps(r1, [
      { type: "addComment", threadId: "t1", c: reply },
      { type: "addComment", threadId: "t1", c: reply },
    ]);
    expect(r.threads[0]!.comments).toHaveLength(2);
  });

  it("setThreadStatus inverts to the previous status", () => {
    const r1 = applyReviewOp(emptyReview("d"), { type: "addThread", t: thread("t1", cmt("c1", "x")) }).layer;
    roundtrip(r1, { type: "setThreadStatus", threadId: "t1", status: "resolved" });
  });
});

describe("ReviewOps commute (order independence)", () => {
  it("two authors' independent adds reach the same layer regardless of order", () => {
    const a: ReviewOp = { type: "addSuggestion", s: { ...sug("sa", "insert"), author: alice } };
    const b: ReviewOp = { type: "addSuggestion", s: { ...sug("sb", "delete"), author: bob } };
    const ab = applyReviewOps(emptyReview("d"), [a, b]);
    const ba = applyReviewOps(emptyReview("d"), [b, a]);
    // order of the suggestions array may differ; compare as sets by id
    expect(new Set(ab.suggestions.map((s) => s.id))).toEqual(new Set(ba.suggestions.map((s) => s.id)));
    expect(ab.suggestions).toHaveLength(2);
  });
});
