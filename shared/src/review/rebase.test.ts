import { describe, expect, it } from "vitest";
import type { Document, Paragraph, SectionProps } from "../model/document";
import { applyOp, type Op } from "../model/ops";
import type { UserInfo } from "../change";
import type { ReviewLayer, Suggestion } from "./model";
import { emptyReview } from "./model";
import { gcReviewLayer, mapReviewAnchors, rebaseReview } from "./rebase";

// Drives the real core applyOp so anchors are rebased through the SAME
// position-mapper the editor uses (the rebaseBookmarks path).
const section = (): SectionProps => ({ pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } });
const CHAR = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA = { align: "left" as const, lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const para = (id: string, text: string): Paragraph => ({ kind: "paragraph", id, revision: 0, runs: [{ text, style: CHAR }], style: PARA });
const docOf = (...ps: Paragraph[]): Document => ({ section: section(), blocks: ps });
const alice: UserInfo = { id: "u", firstName: "A", lastName: "A" };

const pos = (blockId: string, offset: number) => ({ blockId, offset });
const sug = (id: string, s: number, e: number, blockId = "p"): Suggestion => ({
  id,
  kind: "delete",
  anchor: { start: pos(blockId, s), end: pos(blockId, e) },
  author: alice,
  createdAt: 1,
});
const layerWith = (...ss: Suggestion[]): ReviewLayer => ({ ...emptyReview("d"), suggestions: ss });

describe("mapReviewAnchors — anchors travel with edits", () => {
  it("insert before the anchor shifts it right", () => {
    const doc = docOf(para("p", "hello world"));
    const res = applyOp(doc, { type: "insertText", at: pos("p", 0), text: "XX" });
    const l = mapReviewAnchors(layerWith(sug("s", 6, 11)), res.mapPosition);
    expect(l.suggestions[0]!.anchor).toEqual({ start: pos("p", 8), end: pos("p", 13) });
  });

  it("insert inside the anchor grows it", () => {
    const doc = docOf(para("p", "hello world"));
    const res = applyOp(doc, { type: "insertText", at: pos("p", 8), text: "ZZ" }); // inside [6,11)
    const l = mapReviewAnchors(layerWith(sug("s", 6, 11)), res.mapPosition);
    expect(l.suggestions[0]!.anchor).toEqual({ start: pos("p", 6), end: pos("p", 13) });
  });

  it("deleting the whole anchored range collapses it (→ GC'd)", () => {
    const doc = docOf(para("p", "hello world"));
    const res = applyOp(doc, { type: "deleteRange", blockId: "p", start: 6, end: 11 });
    const l = rebaseReview(layerWith(sug("s", 6, 11)), res.mapPosition);
    expect(l.suggestions).toHaveLength(0); // dead record dropped
  });

  it("splitParagraph moves the tail anchor into the new block", () => {
    const doc = docOf(para("p", "hello world"));
    const res = applyOp(doc, { type: "splitParagraph", at: pos("p", 5), newBlockId: "p2" });
    const l = mapReviewAnchors(layerWith(sug("s", 6, 11)), res.mapPosition);
    expect(l.suggestions[0]!.anchor).toEqual({ start: pos("p2", 1), end: pos("p2", 6) });
  });

  it("an unrelated edit on another block leaves the anchor untouched", () => {
    const doc = docOf(para("p", "hello"), para("q", "world"));
    const res = applyOp(doc, { type: "insertText", at: pos("q", 0), text: "ZZ" });
    const l = mapReviewAnchors(layerWith(sug("s", 0, 5)), res.mapPosition);
    expect(l.suggestions[0]!.anchor).toEqual({ start: pos("p", 0), end: pos("p", 5) });
  });
});

describe("gcReviewLayer", () => {
  it("drops collapsed suggestions but keeps point-comment threads", () => {
    const l: ReviewLayer = {
      ...emptyReview("d"),
      suggestions: [sug("dead", 4, 4), sug("live", 0, 3)],
      threads: [{ id: "t", anchor: { start: pos("p", 2), end: pos("p", 2) }, status: "open", comments: [] }],
    };
    const gc = gcReviewLayer(l);
    expect(gc.suggestions.map((s) => s.id)).toEqual(["live"]);
    expect(gc.threads).toHaveLength(1); // point comment survives
  });
});
