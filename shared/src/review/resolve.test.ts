import { describe, expect, it } from "vitest";
import type { Document, Paragraph, SectionProps } from "../model/document";
import { applyOp, type Op } from "../model/ops";
import { textOfRuns } from "../model/text";
import type { UserInfo } from "../change";
import type { ReviewLayer, StructuralChange, Suggestion } from "./model";
import { emptyReview } from "./model";
import { acceptSuggestion, rejectSuggestion, rejectAllSuggestions } from "./resolve";

const section = (): SectionProps => ({ pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } });
const CHAR = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA = { align: "left" as const, lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const para = (id: string, text: string): Paragraph => ({ kind: "paragraph", id, revision: 0, runs: [{ text, style: CHAR }], style: PARA });
const docOf = (...ps: Paragraph[]): Document => ({ section: section(), blocks: ps });
const alice: UserInfo = { id: "u", firstName: "A", lastName: "A" };

const bodyText = (d: Document): string[] =>
  d.blocks.map((b) => (b.kind === "paragraph" ? textOfRuns(b.runs) : `<${b.kind}>`));

/** Mirror intercept: apply the forward op to the live doc, capture its inverse,
 *  and record a structural suggestion. Returns the live (already-applied) doc +
 *  the layer the reviewer acts on. */
function trackStructural(doc: Document, op: Op, blockId: string, id = "s1", createdAt = 1): { doc: Document; review: ReviewLayer } {
  const res = applyOp(doc, op);
  const structural: StructuralChange = { op, inverse: res.inverse, blockId };
  const s: Suggestion = { id, kind: "structural", anchor: { start: { blockId, offset: 0 }, end: { blockId, offset: 0 } }, author: alice, createdAt, structural };
  return { doc: res.doc, review: { ...emptyReview("d"), suggestions: [s] } };
}

const applyAll = (doc: Document, ops: Op[]): Document => ops.reduce((d, op) => applyOp(d, op).doc, doc);

describe("resolve — structural suggestions", () => {
  it("reject a split re-merges the paragraph (restores original structure)", () => {
    const original = docOf(para("p", "hello world"));
    const { doc, review } = trackStructural(
      original,
      { type: "splitParagraph", at: { blockId: "p", offset: 5 }, newBlockId: "p2" },
      "p2",
    );
    expect(bodyText(doc)).toEqual(["hello", " world"]); // split applied to live doc

    const r = rejectSuggestion(doc, review, "s1")!;
    expect(r.ops).toHaveLength(1);
    expect(r.ops[0]!.type).toBe("mergeParagraphs");
    const restored = applyAll(doc, r.ops);
    expect(bodyText(restored)).toEqual(["hello world"]); // back to one paragraph
    expect(r.reviewOps).toEqual([{ type: "removeSuggestion", id: "s1" }]);
  });

  it("accept a split is a pure record drop (text already split)", () => {
    const original = docOf(para("p", "hello world"));
    const { doc, review } = trackStructural(
      original,
      { type: "splitParagraph", at: { blockId: "p", offset: 5 }, newBlockId: "p2" },
      "p2",
    );
    const r = acceptSuggestion(doc, review, "s1")!;
    expect(r.ops).toHaveLength(0); // nothing to do — already applied
    expect(r.reviewOps).toEqual([{ type: "removeSuggestion", id: "s1" }]);
  });

  it("reject a removeBlock re-inserts the block at its original index", () => {
    const original = docOf(para("p", "keep"), para("q", "gone"), para("r", "tail"));
    const { doc, review } = trackStructural(original, { type: "removeBlock", blockId: "q" }, "q");
    expect(bodyText(doc)).toEqual(["keep", "tail"]); // q removed in live doc

    const r = rejectSuggestion(doc, review, "s1")!;
    const restored = applyAll(doc, r.ops);
    expect(bodyText(restored)).toEqual(["keep", "gone", "tail"]); // q is back, in place
  });

  it("reject an insertBlock removes the inserted block", () => {
    const original = docOf(para("p", "a"), para("q", "b"));
    const fresh = para("new", "inserted");
    const { doc, review } = trackStructural(original, { type: "insertBlock", index: 1, block: fresh }, "new");
    expect(bodyText(doc)).toEqual(["a", "inserted", "b"]);

    const r = rejectSuggestion(doc, review, "s1")!;
    const restored = applyAll(doc, r.ops);
    expect(bodyText(restored)).toEqual(["a", "b"]);
  });

  it("rejectAll reverses a split+insert paste newest-first, restoring the baseline", () => {
    // Simulate a 2-block paste decomposed into: split at the caret, then insert a
    // block between the halves. rejectAll must undo the insert THEN the split.
    const original = docOf(para("p", "hello world"));
    const splitOp: Op = { type: "splitParagraph", at: { blockId: "p", offset: 5 }, newBlockId: "p2" };
    const r1 = applyOp(original, splitOp);
    const fresh = para("pasted", "PASTED");
    const insertOp: Op = { type: "insertBlock", index: 1, block: fresh };
    const r2 = applyOp(r1.doc, insertOp);
    const live = r2.doc;
    expect(bodyText(live)).toEqual(["hello", "PASTED", " world"]);

    const review: ReviewLayer = {
      ...emptyReview("d"),
      suggestions: [
        { id: "split", kind: "structural", anchor: { start: { blockId: "p2", offset: 0 }, end: { blockId: "p2", offset: 0 } }, author: alice, createdAt: 1, groupId: "g", structural: { op: splitOp, inverse: r1.inverse, blockId: "p2" } },
        { id: "ins", kind: "structural", anchor: { start: { blockId: "pasted", offset: 0 }, end: { blockId: "pasted", offset: 0 } }, author: alice, createdAt: 2, groupId: "g", structural: { op: insertOp, inverse: r2.inverse, blockId: "pasted" } },
      ],
    };

    const res = rejectAllSuggestions(live, review);
    const restored = applyAll(live, res.ops);
    expect(bodyText(restored)).toEqual(["hello world"]); // fully back to baseline
  });
});
