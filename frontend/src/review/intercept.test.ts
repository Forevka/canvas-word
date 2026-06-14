import { describe, expect, it } from "vitest";
import {
  applyStylePatchToRuns,
  emptyReview,
  type CharStyle,
  type Document,
  type Paragraph,
  type ReviewLayer,
  type SectionProps,
  type Suggestion,
  type UserInfo,
} from "@cw/shared";
import type { Transaction } from "../editor/state";
import { intercept } from "./intercept";

const section = (): SectionProps => ({ pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } });
const CHAR: CharStyle = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA = { align: "left" as const, lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const para = (id: string, text: string): Paragraph => ({ kind: "paragraph", id, revision: 0, runs: [{ text, style: CHAR }], style: PARA });
const docOf = (...ps: Paragraph[]): Document => ({ section: section(), blocks: ps });
const alice: UserInfo = { id: "u-alice", firstName: "Alice", lastName: "A" };
const bob: UserInfo = { id: "u-bob", firstName: "Bob", lastName: "B" };
const caret = (blockId: string, offset: number) => ({ anchor: { blockId, offset }, focus: { blockId, offset } });
const tr = (ops: Transaction["ops"], blockId: string, offset: number, origin: Transaction["origin"] = "typing"): Transaction => ({
  ops,
  selectionAfter: caret(blockId, offset),
  origin,
});
const insertRecord = (id: string, blockId: string, s: number, e: number, author = alice): Suggestion => ({
  id,
  kind: "insert",
  anchor: { start: { blockId, offset: s }, end: { blockId, offset: e } },
  author,
  createdAt: 1,
});
const withSuggestions = (...ss: Suggestion[]): ReviewLayer => ({ ...emptyReview("d"), suggestions: ss });

describe("intercept — insertions", () => {
  it("fresh insertion point creates one insert suggestion in post-insert coords", () => {
    const doc = docOf(para("p", "hello"));
    const t = tr([{ type: "insertText", at: { blockId: "p", offset: 5 }, text: "XYZ" }], "p", 8);
    const r = intercept(t, emptyReview("d"), doc, alice, 100);
    expect(r.core.ops).toEqual(t.ops); // text still really inserted
    expect(r.reviewOps).toHaveLength(1);
    const op = r.reviewOps[0]!;
    expect(op.type).toBe("addSuggestion");
    if (op.type === "addSuggestion") {
      expect(op.s.kind).toBe("insert");
      expect(op.s.anchor).toEqual({ start: { blockId: "p", offset: 5 }, end: { blockId: "p", offset: 8 } });
      expect(op.s.author.id).toBe("u-alice");
    }
  });

  it("typing at the end of the author's own insertion adds NO new record (rebase grows it)", () => {
    const doc = docOf(para("p", "helloXYZ"));
    const layer = withSuggestions(insertRecord("s1", "p", 5, 8)); // 'XYZ' already an insert record
    const t = tr([{ type: "insertText", at: { blockId: "p", offset: 8 }, text: "Q" }], "p", 9);
    const r = intercept(t, layer, doc, alice, 100);
    expect(r.core.ops).toEqual(t.ops);
    expect(r.reviewOps).toHaveLength(0);
  });

  it("typing into the middle of someone else's insertion still creates the author's own record", () => {
    const doc = docOf(para("p", "helloXYZ"));
    const layer = withSuggestions(insertRecord("s1", "p", 5, 8, bob));
    const t = tr([{ type: "insertText", at: { blockId: "p", offset: 6 }, text: "Q" }], "p", 7);
    const r = intercept(t, layer, doc, alice, 100);
    expect(r.reviewOps).toHaveLength(1); // not Alice's insert → new record
  });
});

describe("intercept — deletions", () => {
  it("deleting original text is non-destructive: drops the core delete, marks deleted", () => {
    const doc = docOf(para("p", "hello world"));
    const t = tr([{ type: "deleteRange", blockId: "p", start: 0, end: 5 }], "p", 0);
    const r = intercept(t, emptyReview("d"), doc, alice, 100);
    expect(r.core.ops).toHaveLength(0); // text stays
    expect(r.core.selectionAfter).toEqual(caret("p", 0));
    expect(r.reviewOps).toHaveLength(1);
    const op = r.reviewOps[0]!;
    if (op.type === "addSuggestion") {
      expect(op.s.kind).toBe("delete");
      expect(op.s.anchor).toEqual({ start: { blockId: "p", offset: 0 }, end: { blockId: "p", offset: 5 } });
    }
  });

  it("deleting the author's OWN pending insertion is a real delete (no record; rebase shrinks it)", () => {
    const doc = docOf(para("p", "helloXYZ"));
    const layer = withSuggestions(insertRecord("s1", "p", 5, 8)); // alice inserted 'XYZ'
    const t = tr([{ type: "deleteRange", blockId: "p", start: 5, end: 8 }], "p", 5);
    const r = intercept(t, layer, doc, alice, 100);
    expect(r.core.ops).toEqual(t.ops); // really deleted
    expect(r.reviewOps).toHaveLength(0);
  });

  it("a mixed delete (own insert + original) marks the whole range deleted (V1 rule)", () => {
    const doc = docOf(para("p", "helloXYZworld"));
    const layer = withSuggestions(insertRecord("s1", "p", 5, 8));
    const t = tr([{ type: "deleteRange", blockId: "p", start: 5, end: 13 }], "p", 5); // covers XYZ + 'world'
    const r = intercept(t, layer, doc, alice, 100);
    expect(r.core.ops).toHaveLength(0); // not fully covered → non-destructive
    expect(r.reviewOps).toHaveLength(1);
  });
});

describe("intercept — format", () => {
  it("a style-only setRuns becomes a format suggestion with patch + inverse", () => {
    const doc = docOf(para("p", "hello world"));
    const newRuns = applyStylePatchToRuns(doc.blocks[0]!.kind === "paragraph" ? (doc.blocks[0] as Paragraph).runs : [], 0, 5, { bold: true });
    const t = tr([{ type: "setRuns", blockId: "p", runs: newRuns }], "p", 5, "command");
    const r = intercept(t, emptyReview("d"), doc, alice, 100);
    expect(r.core.ops).toEqual(t.ops); // restyle applied as-is
    expect(r.reviewOps).toHaveLength(1);
    const op = r.reviewOps[0]!;
    if (op.type === "addSuggestion") {
      expect(op.s.kind).toBe("format");
      expect(op.s.anchor).toEqual({ start: { blockId: "p", offset: 0 }, end: { blockId: "p", offset: 5 } });
      expect(op.s.patch).toEqual({ bold: true });
      expect(op.s.inverse).toEqual({ bold: false });
    }
  });
});

describe("intercept — pass-through (untracked V1 boundary)", () => {
  it("structural ops pass through untracked", () => {
    const doc = docOf(para("p", "hello"));
    const t = tr([{ type: "splitParagraph", at: { blockId: "p", offset: 2 }, newBlockId: "p2" }], "p2", 0);
    const r = intercept(t, emptyReview("d"), doc, alice, 100);
    expect(r.core).toBe(t);
    expect(r.reviewOps).toHaveLength(0);
  });

  it("multi-op transactions pass through untracked", () => {
    const doc = docOf(para("p", "hello"), para("q", "world"));
    const t = tr(
      [
        { type: "deleteRange", blockId: "p", start: 2, end: 5 },
        { type: "deleteRange", blockId: "q", start: 0, end: 2 },
      ],
      "p",
      2,
    );
    const r = intercept(t, emptyReview("d"), doc, alice, 100);
    expect(r.core).toBe(t);
    expect(r.reviewOps).toHaveLength(0);
  });
});
