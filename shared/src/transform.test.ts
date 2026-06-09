import { describe, expect, it } from "vitest";
import type { Document, Paragraph, SectionProps } from "./model/document";
import { applyOp, type Op } from "./model/ops";
import { textOfRuns } from "./model/text";
import { transformOp } from "./transform";

const section = (): SectionProps => ({ pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } });
const CHAR = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA = { align: "left" as const, lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const para = (id: string, text: string): Paragraph => ({ kind: "paragraph", id, revision: 0, runs: [{ text, style: CHAR }], style: PARA });
const docOf = (...ps: Paragraph[]): Document => ({ section: section(), blocks: ps });

const ins = (blockId: string, offset: number, text: string): Op => ({ type: "insertText", at: { blockId, offset }, text });
const del = (blockId: string, start: number, end: number): Op => ({ type: "deleteRange", blockId, start, end });
const split = (blockId: string, offset: number, newBlockId: string): Op => ({ type: "splitParagraph", at: { blockId, offset }, newBlockId });

const apply = (d: Document, op: Op): Document => applyOp(d, op).doc;
const applyMany = (d: Document, ops: Op[]): Document => ops.reduce(apply, d);
const docText = (d: Document): string =>
  d.blocks.map((b) => (b.kind === "paragraph" ? textOfRuns(b.runs) : `[${b.kind}]`)).join("⏎");

// The TP1 convergence property: two concurrent ops a (winner/left) and b
// (loser/right) reach the SAME state regardless of which the server orders first.
const converges = (s: Document, a: Op, b: Op): { viaA: string; viaB: string } => ({
  viaA: docText(applyMany(apply(s, a), transformOp(b, a, "right"))), // a first, b yields
  viaB: docText(applyMany(apply(s, b), transformOp(a, b, "left"))), // b first, a keeps
});

describe("transformOp — convergence (TP1)", () => {
  it("concurrent inserts at different offsets", () => {
    const s = docOf(para("p", "abcdef"));
    const { viaA, viaB } = converges(s, ins("p", 1, "X"), ins("p", 4, "Y"));
    expect(viaA).toBe(viaB);
    expect(viaA).toBe("aXbcdYef");
  });

  it("concurrent inserts at the SAME offset (deterministic tie-break)", () => {
    const s = docOf(para("p", "abcdef"));
    const { viaA, viaB } = converges(s, ins("p", 3, "X"), ins("p", 3, "Y"));
    expect(viaA).toBe(viaB);
    expect(viaA).toBe("abcXYdef"); // winner (a) before loser (b)
  });

  it("insert vs delete", () => {
    const s = docOf(para("p", "abcdef"));
    const { viaA, viaB } = converges(s, ins("p", 2, "X"), del("p", 1, 4));
    expect(viaA).toBe(viaB); // "aXef"
    expect(viaA).toBe("aXef");
  });

  it("delete vs insert inside the deleted range", () => {
    const s = docOf(para("p", "abcdef"));
    const { viaA, viaB } = converges(s, del("p", 1, 4), ins("p", 2, "X"));
    expect(viaA).toBe(viaB);
  });

  it("overlapping deletes", () => {
    const s = docOf(para("p", "abcdef"));
    const { viaA, viaB } = converges(s, del("p", 1, 4), del("p", 2, 5));
    expect(viaA).toBe(viaB);
    expect(viaA).toBe("af");
  });

  it("ops on different paragraphs are independent", () => {
    const s = docOf(para("p1", "abc"), para("p2", "xyz"));
    const { viaA, viaB } = converges(s, ins("p1", 1, "A"), ins("p2", 1, "B"));
    expect(viaA).toBe(viaB);
    expect(viaA).toBe("aAbc⏎xByz");
  });

  it("insert after a concurrent split follows into the new paragraph", () => {
    const s = docOf(para("p", "abcdef"));
    // a = split at 3 (winner), b = insert "X" at 5 (after split → new block)
    const { viaA, viaB } = converges(s, split("p", 3, "p-new"), ins("p", 5, "X"));
    expect(viaA).toBe(viaB);
    expect(viaA).toBe("abc⏎deXf");
  });
});

describe("transformOp — coarse ops are last-writer-wins (identity transform)", () => {
  it("setParaStyle is unchanged by a concurrent op", () => {
    const a: Op = { type: "setParaStyle", blockId: "p", patch: { align: "center" } };
    const b = ins("p", 0, "X");
    expect(transformOp(a, b, "right")).toEqual([a]); // no offset surgery for coarse ops
  });
});
