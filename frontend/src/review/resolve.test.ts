import { describe, expect, it } from "vitest";
import {
  applyOp,
  applyStylePatchToRuns,
  emptyReview,
  textOfRuns,
  type CharStyle,
  type Document,
  type Op,
  type Paragraph,
  type ReviewLayer,
  type SectionProps,
  type Suggestion,
  type UserInfo,
} from "@cw/shared";
import { acceptSuggestion, rejectSuggestion, acceptAllSuggestions, rejectAllSuggestions, bakeReview } from "@cw/shared";

const section = (): SectionProps => ({ pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } });
const CHAR: CharStyle = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA = { align: "left" as const, lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const para = (id: string, text: string): Paragraph => ({ kind: "paragraph", id, revision: 0, runs: [{ text, style: CHAR }], style: PARA });
const docOf = (...ps: Paragraph[]): Document => ({ section: section(), blocks: ps });
const alice: UserInfo = { id: "u", firstName: "A", lastName: "A" };
const sug = (id: string, kind: Suggestion["kind"], s: number, e: number, extra?: Partial<Suggestion>): Suggestion => ({
  id,
  kind,
  anchor: { start: { blockId: "p", offset: s }, end: { blockId: "p", offset: e } },
  author: alice,
  createdAt: 1,
  ...extra,
});
const layerOf = (...ss: Suggestion[]): ReviewLayer => ({ ...emptyReview("d"), suggestions: ss });
const docText = (d: Document): string => textOfRuns((d.blocks[0] as Paragraph).runs);
const applyAll = (d: Document, ops: Op[]): Document => ops.reduce((acc, op) => applyOp(acc, op).doc, d);

describe("accept / reject single suggestions", () => {
  it("accept insert keeps text (no core op), drops the record", () => {
    const r = acceptSuggestion(docOf(para("p", "abc")), layerOf(sug("s", "insert", 0, 1)), "s")!;
    expect(r.ops).toHaveLength(0);
    expect(r.reviewOps).toEqual([{ type: "removeSuggestion", id: "s" }]);
  });

  it("reject insert deletes the inserted text", () => {
    const doc = docOf(para("p", "aXbc"));
    const r = rejectSuggestion(doc, layerOf(sug("s", "insert", 1, 2)), "s")!;
    expect(r.ops).toEqual([{ type: "deleteRange", blockId: "p", start: 1, end: 2 }]);
    expect(docText(applyAll(doc, r.ops))).toBe("abc");
  });

  it("accept delete removes the marked text", () => {
    const doc = docOf(para("p", "aXbc"));
    const r = acceptSuggestion(doc, layerOf(sug("s", "delete", 1, 2)), "s")!;
    expect(docText(applyAll(doc, r.ops))).toBe("abc");
  });

  it("reject delete keeps text (no core op)", () => {
    const r = rejectSuggestion(docOf(para("p", "aXbc")), layerOf(sug("s", "delete", 1, 2)), "s")!;
    expect(r.ops).toHaveLength(0);
  });

  it("reject format re-applies the inverse patch", () => {
    const doc = docOf(para("p", "hello"));
    const bolded = { ...doc, blocks: [{ ...(doc.blocks[0] as Paragraph), runs: applyStylePatchToRuns((doc.blocks[0] as Paragraph).runs, 0, 5, { bold: true }) }] };
    const r = rejectSuggestion(bolded, layerOf(sug("s", "format", 0, 5, { patch: { bold: true }, inverse: { bold: false } })), "s")!;
    const out = applyAll(bolded, r.ops).blocks[0] as Paragraph;
    expect(out.runs.every((run) => run.style.bold === false)).toBe(true);
  });
});

// Live doc "AXBYC": X@[1,2) is a tracked insert, Y@[3,4) a tracked delete.
const scenario = (): { doc: Document; review: ReviewLayer } => ({
  doc: docOf(para("p", "AXBYC")),
  review: layerOf(sug("ins", "insert", 1, 2), sug("del", "delete", 3, 4)),
});

describe("bake (export folding)", () => {
  it("reject-all = original baseline (insertions removed, deletions kept)", () => {
    const { doc, review } = scenario();
    expect(docText(bakeReview(doc, review, "reject"))).toBe("ABYC");
  });

  it("accept-all = final text (insertions kept, deletions removed)", () => {
    const { doc, review } = scenario();
    expect(docText(bakeReview(doc, review, "accept"))).toBe("AXBC");
  });

  it("acceptAll/rejectAll resolutions apply cleanly back-to-front", () => {
    const { doc, review } = scenario();
    expect(docText(applyAll(doc, rejectAllSuggestions(doc, review).ops))).toBe("ABYC");
    expect(docText(applyAll(doc, acceptAllSuggestions(doc, review).ops))).toBe("AXBC");
  });
});
