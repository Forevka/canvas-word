import { describe, expect, it } from "vitest";
import type { Document, Paragraph, SectionProps } from "../model/document";
import { textOfRuns } from "../model/text";
import { assembleMerge, merge3, type MergeChoice } from "./merge3";

const section = (): SectionProps => ({ pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } });
const CHAR = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA = { align: "left" as const, lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const para = (id: string, text: string): Paragraph => ({ kind: "paragraph", id, revision: 0, runs: [{ text, style: CHAR }], style: PARA });
const docOf = (...t: string[]): Document => ({ section: section(), blocks: t.map((s, i) => para(`b${i}`, s)) });

const body = (d: Document): string[] => d.blocks.map((b) => (b.kind === "paragraph" ? textOfRuns(b.runs) : `<${b.kind}>`));
// Assemble from a SINGLE merge result (conflict ids are minted per merge3 call).
const asm = (res: ReturnType<typeof merge3>, choices: Record<string, MergeChoice> = {}): string[] => body(assembleMerge(res, choices));
const assembled = (base: Document, mine: Document, theirs: Document, choices: Record<string, MergeChoice> = {}): string[] =>
  asm(merge3(base, mine, theirs), choices);

describe("merge3 — 3-way block merge", () => {
  it("no changes → base", () => {
    const base = docOf("a", "b", "c");
    const res = merge3(base, docOf("a", "b", "c"), docOf("a", "b", "c"));
    expect(res.stats.conflicts).toBe(0);
    expect(res.stats.autoMerged).toBe(0);
    expect(assembled(base, docOf("a", "b", "c"), docOf("a", "b", "c"))).toEqual(["a", "b", "c"]);
  });

  it("change on mine only → takes mine", () => {
    const base = docOf("a", "b", "c");
    const mine = docOf("a", "B2", "c");
    const theirs = docOf("a", "b", "c");
    const res = merge3(base, mine, theirs);
    expect(res.stats.conflicts).toBe(0);
    expect(assembled(base, mine, theirs)).toEqual(["a", "B2", "c"]);
  });

  it("change on theirs only → takes theirs", () => {
    const base = docOf("a", "b", "c");
    expect(assembled(base, docOf("a", "b", "c"), docOf("a", "b", "C3"))).toEqual(["a", "b", "C3"]);
  });

  it("same change on both sides → auto-merges once", () => {
    const base = docOf("a", "b", "c");
    const both = docOf("a", "X", "c");
    const res = merge3(base, both, both);
    expect(res.stats.conflicts).toBe(0);
    expect(res.stats.autoMerged).toBe(1);
    expect(assembled(base, both, both)).toEqual(["a", "X", "c"]);
  });

  it("non-overlapping changes on each side → both auto-merge", () => {
    const base = docOf("a", "b", "c", "d");
    const mine = docOf("A1", "b", "c", "d"); // changed first
    const theirs = docOf("a", "b", "c", "D4"); // changed last
    const res = merge3(base, mine, theirs);
    expect(res.stats.conflicts).toBe(0);
    expect(assembled(base, mine, theirs)).toEqual(["A1", "b", "c", "D4"]);
  });

  it("diverging change on the same region → conflict, resolvable by side", () => {
    const base = docOf("a", "b", "c");
    const mine = docOf("a", "MINE", "c");
    const theirs = docOf("a", "THEIRS", "c");
    const res = merge3(base, mine, theirs);
    expect(res.stats.conflicts).toBe(1);
    const id = res.conflicts[0]!.id!;
    expect(asm(res)).toEqual(["a", "MINE", "c"]); // default = mine
    expect(asm(res, { [id]: "theirs" })).toEqual(["a", "THEIRS", "c"]);
    expect(asm(res, { [id]: "both" })).toEqual(["a", "MINE", "THEIRS", "c"]);
    expect(asm(res, { [id]: "base" })).toEqual(["a", "b", "c"]);
  });

  it("insertion on mine + different insertion on theirs at the same spot → conflict", () => {
    const base = docOf("a", "b");
    const mine = docOf("a", "m1", "m2", "b");
    const theirs = docOf("a", "t1", "b");
    const res = merge3(base, mine, theirs);
    expect(res.stats.conflicts).toBe(1);
    const id = res.conflicts[0]!.id!;
    expect(asm(res, { [id]: "mine" })).toEqual(["a", "m1", "m2", "b"]);
    expect(asm(res, { [id]: "theirs" })).toEqual(["a", "t1", "b"]);
  });

  it("deletion on one side, edit on the other → conflict", () => {
    const base = docOf("a", "b", "c");
    const mine = docOf("a", "c"); // deleted b
    const theirs = docOf("a", "B!", "c"); // edited b
    const res = merge3(base, mine, theirs);
    expect(res.stats.conflicts).toBe(1);
    const id = res.conflicts[0]!.id!;
    expect(asm(res, { [id]: "mine" })).toEqual(["a", "c"]);
    expect(asm(res, { [id]: "theirs" })).toEqual(["a", "B!", "c"]);
  });

  it("appends on both sides merge without conflict when disjoint", () => {
    const base = docOf("a");
    const mine = docOf("a", "m"); // appended m
    const theirs = docOf("a"); // unchanged
    expect(assembled(base, mine, theirs)).toEqual(["a", "m"]);
  });
});
