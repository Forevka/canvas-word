import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, Run, SectionProps } from "../model/document";
import { normalizeRuns } from "../model/ops";
import { bakeReview } from "../review/bake";
import { diffDocuments } from "./diffDocuments";

const section = (): SectionProps => ({ pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } });
const CHAR: CharStyle = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA: ParaStyle = { align: "left", lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };

const run = (text: string, patch: Partial<CharStyle> = {}): Run => ({ text, style: { ...CHAR, ...patch } });
const para = (id: string, text: string, paraPatch: Partial<ParaStyle> = {}): Paragraph => ({ kind: "paragraph", id, revision: 0, runs: [run(text)], style: { ...PARA, ...paraPatch } });
const paraR = (id: string, runs: Run[], paraPatch: Partial<ParaStyle> = {}): Paragraph => ({ kind: "paragraph", id, revision: 0, runs, style: { ...PARA, ...paraPatch } });
const docOf = (...ps: Paragraph[]): Document => ({ section: section(), blocks: ps });

/** A comparable snapshot of a doc's body — ignores block ids/revisions and
 *  coalesces equal-styled runs, so two docs that render identically compare equal. */
const snap = (d: Document): unknown =>
  d.blocks.map((b) =>
    b.kind === "paragraph"
      ? { k: "p", style: b.style, runs: normalizeRuns(b.runs, b.runs[0]?.style ?? CHAR).map((r) => ({ t: r.text, s: r.style })) }
      : { k: b.kind },
  );

/** The core correctness oracle: reject → original, accept → revised. */
const expectRoundTrip = (original: Document, revised: Document): ReturnType<typeof diffDocuments> => {
  const res = diffDocuments(original, revised, { docId: "d" });
  expect(snap(bakeReview(res.mergedDoc, res.review, "reject"))).toEqual(snap(original));
  expect(snap(bakeReview(res.mergedDoc, res.review, "accept"))).toEqual(snap(revised));
  return res;
};

describe("diffDocuments — round-trip oracle", () => {
  it("identical documents produce no suggestions", () => {
    const a = docOf(para("p", "hello world"), para("q", "second"));
    const res = expectRoundTrip(a, docOf(para("x", "hello world"), para("y", "second")));
    expect(res.review.suggestions).toHaveLength(0);
  });

  it("pure text insertion (a word added)", () => {
    const a = docOf(para("p", "the quick fox"));
    const b = docOf(para("p2", "the quick brown fox"));
    const res = expectRoundTrip(a, b);
    expect(res.stats.insertions).toBeGreaterThan(0);
    expect(res.stats.deletions).toBe(0);
    expect(res.review.suggestions.some((s) => s.kind === "insert")).toBe(true);
  });

  it("pure text deletion (a word removed)", () => {
    const a = docOf(para("p", "the quick brown fox"));
    const b = docOf(para("p2", "the quick fox"));
    const res = expectRoundTrip(a, b);
    expect(res.stats.deletions).toBeGreaterThan(0);
    expect(res.review.suggestions.some((s) => s.kind === "delete")).toBe(true);
  });

  it("word replacement", () => {
    expectRoundTrip(docOf(para("p", "the quick brown fox")), docOf(para("p2", "the slow brown fox")));
  });

  it("character-format-only change (a word bolded)", () => {
    const a = docOf(paraR("p", [run("make this "), run("word"), run(" bold")]));
    const b = docOf(paraR("p2", [run("make this "), run("word", { bold: true }), run(" bold")]));
    const res = expectRoundTrip(a, b);
    expect(res.stats.formatChanges).toBeGreaterThan(0);
    expect(res.review.suggestions.some((s) => s.kind === "format")).toBe(true);
  });

  it("paragraph-style-only change (alignment)", () => {
    const a = docOf(para("p", "centered later", { align: "left" }));
    const b = docOf(para("p2", "centered later", { align: "center" }));
    const res = expectRoundTrip(a, b);
    expect(res.stats.paraStyleChanges).toBe(1);
    expect(res.review.suggestions.some((s) => s.kind === "structural")).toBe(true);
  });

  it("whole block inserted (revised has an extra paragraph)", () => {
    const a = docOf(para("p", "one"), para("q", "three"));
    const b = docOf(para("x", "one"), para("y", "two"), para("z", "three"));
    const res = expectRoundTrip(a, b);
    expect(res.stats.blockInserts).toBe(1);
  });

  it("whole block deleted (original has an extra paragraph)", () => {
    const a = docOf(para("p", "one"), para("q", "two"), para("r", "three"));
    const b = docOf(para("x", "one"), para("z", "three"));
    const res = expectRoundTrip(a, b);
    expect(res.stats.blockDeletes).toBe(1);
  });

  it("combined: insert word, delete word, add paragraph, bold a word", () => {
    const a = docOf(paraR("p", [run("alpha "), run("beta"), run(" gamma")]), para("q", "keep me"), para("r", "remove me"));
    const b = docOf(
      paraR("p2", [run("alpha inserted "), run("beta", { italic: true }), run("")]),
      para("q2", "keep me"),
      para("n", "brand new"),
    );
    expectRoundTrip(a, b);
  });

  it("multi-paragraph rewrite (block replace)", () => {
    const a = docOf(para("p", "totally different original text here"));
    const b = docOf(para("p2", "completely new revised content"));
    expectRoundTrip(a, b);
  });

  it("empty paragraph gains text", () => {
    expectRoundTrip(docOf(para("p", "")), docOf(para("p2", "now has content")));
  });

  it("char granularity also round-trips", () => {
    const a = docOf(para("p", "kitten"));
    const b = docOf(para("p2", "sitting"));
    const res = diffDocuments(a, b, { granularity: "char", docId: "d" });
    expect(snap(bakeReview(res.mergedDoc, res.review, "reject"))).toEqual(snap(a));
    expect(snap(bakeReview(res.mergedDoc, res.review, "accept"))).toEqual(snap(b));
  });

  it("merged document holds BOTH sides at once (union invariant)", () => {
    const a = docOf(para("p", "the brown fox"));
    const b = docOf(para("p2", "the red fox"));
    const { mergedDoc } = diffDocuments(a, b, { docId: "d" });
    const text = (mergedDoc.blocks[0] as Paragraph).runs.map((r) => r.text).join("");
    expect(text).toContain("brown"); // original word still present (struck)
    expect(text).toContain("red"); // revised word spliced in (underlined)
  });
});
