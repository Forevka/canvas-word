import { describe, expect, it } from "vitest";
import type { Change, UserInfo } from "../change";
import type { Document, Paragraph, SectionProps } from "../model/document";
import type { Op } from "../model/ops";
import { SNAPSHOT_VERSION, type SerializedDocument } from "../persist/serialize";
import type { CommentThread, Suggestion } from "./model";
import type { ReviewOpEnvelope } from "./ops";
import { rehydrateReview } from "./rehydrate";

const section = (): SectionProps => ({ pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } });
const CHAR = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA = { align: "left" as const, lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const para = (id: string, text: string): Paragraph => ({ kind: "paragraph", id, revision: 0, runs: [{ text, style: CHAR }], style: PARA });
const docOf = (...ps: Paragraph[]): Document => ({ section: section(), blocks: ps });
const snap = (doc: Document): SerializedDocument => ({ version: SNAPSHOT_VERSION, doc });
const alice: UserInfo = { id: "u", firstName: "A", lastName: "A" };

const change = (seq: number, ops: Op[]): Change => ({ id: `c${seq}`, docId: "d", baseVersion: seq, seq, siteId: "s", origin: "command", ts: 1, ops });
const pos = (o: number) => ({ blockId: "p", offset: o });
const thread = (id: string, s: number, e: number): CommentThread => ({ id, anchor: { start: pos(s), end: pos(e) }, status: "open", comments: [{ id: id + "c", author: alice, body: [{ text: "note", style: CHAR }], createdAt: 1 }] });
const sug = (id: string, s: number, e: number): Suggestion => ({ id, kind: "delete", anchor: { start: pos(s), end: pos(e) }, author: alice, createdAt: 1 });
const env = (op: ReviewOpEnvelope["op"], dependsOnSeq: number): ReviewOpEnvelope => ({ op, dependsOnSeq, author: alice });

describe("rehydrateReview — fast-forward anchors to head", () => {
  it("maps a comment authored BEFORE later edits forward to head", () => {
    const base = snap(docOf(para("p", "Hello world")));
    // seq0 inserts "XX" at offset 0 → head text "XXHello world".
    const changes = [change(0, [{ type: "insertText", at: pos(0), text: "XX" }])];
    // comment on [0,5] authored at version 0 (before the insert).
    const layer = rehydrateReview("d", base, 0, changes, [env({ type: "addThread", t: thread("t1", 0, 5) }, 0)]);
    expect(layer.threads).toHaveLength(1);
    expect(layer.threads[0]!.anchor).toEqual({ start: pos(2), end: pos(7) }); // shifted +2
    expect(layer.baseVersion).toBe(1);
  });

  it("leaves an anchor authored AFTER the edit untouched", () => {
    const base = snap(docOf(para("p", "Hello world")));
    const changes = [change(0, [{ type: "insertText", at: pos(0), text: "XX" }])];
    // comment authored at version 1 (after the insert) with head-coords [2,7].
    const layer = rehydrateReview("d", base, 0, changes, [env({ type: "addThread", t: thread("t1", 2, 7) }, 1)]);
    expect(layer.threads[0]!.anchor).toEqual({ start: pos(2), end: pos(7) });
  });

  it("folds add+remove so a resolved/rejected suggestion does not resurface", () => {
    const base = snap(docOf(para("p", "Hello world")));
    const layer = rehydrateReview("d", base, 0, [], [
      env({ type: "addSuggestion", s: sug("s1", 0, 5) }, 0),
      env({ type: "removeSuggestion", id: "s1" }, 0),
    ]);
    expect(layer.suggestions).toHaveLength(0);
  });

  it("drops a suggestion whose text was deleted by a later edit (collapsed → GC)", () => {
    const base = snap(docOf(para("p", "Hello world")));
    // delete "Hello " (offsets 0..6) at seq0.
    const changes = [change(0, [{ type: "deleteRange", blockId: "p", start: 0, end: 6 }])];
    const layer = rehydrateReview("d", base, 0, changes, [env({ type: "addSuggestion", s: sug("s1", 0, 5) }, 0)]);
    expect(layer.suggestions).toHaveLength(0);
  });
});
