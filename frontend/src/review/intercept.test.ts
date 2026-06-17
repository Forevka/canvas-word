import { describe, expect, it } from "vitest";
import {
  applyOp,
  applyReviewOps,
  applyStylePatchToRuns,
  emptyReview,
  rejectAllSuggestions,
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

describe("intercept — structural suggestions", () => {
  it("splitParagraph stays in core and yields one structural suggestion carrying op + inverse", () => {
    const doc = docOf(para("p", "hello"));
    const t = tr([{ type: "splitParagraph", at: { blockId: "p", offset: 2 }, newBlockId: "p2" }], "p2", 0);
    const r = intercept(t, emptyReview("d"), doc, alice, 100);
    expect(r.core.ops).toEqual(t.ops); // split really applied (doc stays live)
    expect(r.reviewOps).toHaveLength(1);
    const op = r.reviewOps[0]!;
    expect(op.type).toBe("addSuggestion");
    if (op.type === "addSuggestion") {
      expect(op.s.kind).toBe("structural");
      expect(op.s.structural?.op).toEqual(t.ops[0]);
      expect(op.s.structural?.blockId).toBe("p2"); // hangs on the created block
      expect(op.s.structural?.inverse).toEqual({ type: "mergeParagraphs", firstBlockId: "p" });
    }
  });

  it("removeBlock keeps the text in core (no destructive op dropped) and tracks it", () => {
    const doc = docOf(para("p", "keep"), para("q", "gone"));
    const t = tr([{ type: "removeBlock", blockId: "q" }], "p", 0);
    const r = intercept(t, emptyReview("d"), doc, alice, 100);
    // The remove still runs in core; rejecting re-inserts via the stored inverse.
    expect(r.core.ops).toEqual(t.ops);
    const op = r.reviewOps[0]!;
    if (op.type === "addSuggestion") {
      expect(op.s.kind).toBe("structural");
      expect(op.s.structural?.blockId).toBe("q");
      expect(op.s.structural?.inverse.type).toBe("insertBlock");
    }
  });

  it("a structural op on a missing block passes through untracked", () => {
    const doc = docOf(para("p", "hello"));
    const t = tr([{ type: "mergeParagraphs", firstBlockId: "nope" }], "p", 0);
    const r = intercept(t, emptyReview("d"), doc, alice, 100);
    expect(r.core).toBe(t);
    expect(r.reviewOps).toHaveLength(0);
  });
});

// Mirror the runtime commit: apply the rewritten core ops to the doc (the
// review layer is review-op-only here, so no anchor rebase is needed beyond what
// intercept already baked in), then apply the emitted review ops. Returns the
// committed doc + layer the reviewer acts on.
const commitInterceptResult = (doc: Document, r: ReturnType<typeof intercept>): { doc: Document; review: ReviewLayer } => {
  let d = doc;
  for (const op of r.core.ops) d = applyOp(d, op).doc;
  const review = applyReviewOps(emptyReview("d"), r.reviewOps);
  return { doc: d, review };
};
const bodyText = (d: Document): string[] =>
  d.blocks.map((b) => (b.kind === "paragraph" ? textOfRuns(b.runs) : `<${b.kind}>`));
const run = (text: string): Paragraph["runs"][number] => ({ text, style: CHAR });

describe("intercept — multi-op transactions (paste / cross-paragraph delete)", () => {
  it("a multi-block paste applies structure+text to core and emits grouped suggestions", () => {
    const doc = docOf(para("p", "hello world"));
    // Paste "FOO\nBAR" at offset 5: split, insert into the new tail, then more.
    const t = tr(
      [
        { type: "splitParagraph", at: { blockId: "p", offset: 5 }, newBlockId: "p2" },
        { type: "insertRuns", at: { blockId: "p", offset: 5 }, runs: [run("FOO")] },
        { type: "insertRuns", at: { blockId: "p2", offset: 0 }, runs: [run("BAR")] },
      ],
      "p2",
      3,
    );
    const r = intercept(t, emptyReview("d"), doc, alice, 100);
    expect(r.core.ops.length).toBe(3); // all structure+text really applied
    expect(r.reviewOps.length).toBe(3); // one structural + two inserts
    const groups = new Set(r.reviewOps.map((o) => (o.type === "addSuggestion" ? o.s.groupId : undefined)));
    expect(groups.size).toBe(1); // one shared groupId — accept/reject as a unit
    expect([...groups][0]).toBeTruthy();
  });

  it("paste then reject-all restores the original single paragraph", () => {
    const original = docOf(para("p", "hello world"));
    const t = tr(
      [
        { type: "splitParagraph", at: { blockId: "p", offset: 5 }, newBlockId: "p2" },
        { type: "insertRuns", at: { blockId: "p", offset: 5 }, runs: [run("FOO")] },
        { type: "insertRuns", at: { blockId: "p2", offset: 0 }, runs: [run("BAR")] },
      ],
      "p2",
      3,
    );
    const r = intercept(t, emptyReview("d"), original, alice, 100);
    const { doc, review } = commitInterceptResult(original, r);
    expect(bodyText(doc)).toEqual(["helloFOO", "BAR world"]); // pasted live

    const res = rejectAllSuggestions(doc, review);
    let d = doc;
    for (const op of res.ops) d = applyOp(d, op).doc;
    expect(bodyText(d)).toEqual(["hello world"]); // fully reverted
  });

  it("cross-paragraph delete is tracked: text stays, reject-all is a no-op leaving it", () => {
    // Select from p:2 to q:2 and delete: deleteRange p[2..5], deleteRange q[0..2],
    // mergeParagraphs(p). Tracked → text kept, paragraphs unmerged in core.
    const original = docOf(para("p", "hello"), para("q", "world"));
    const t = tr(
      [
        { type: "deleteRange", blockId: "p", start: 2, end: 5 },
        { type: "deleteRange", blockId: "q", start: 0, end: 2 },
        { type: "mergeParagraphs", firstBlockId: "p" } as Op,
      ],
      "p",
      2,
    );
    const r = intercept(t, emptyReview("d"), original, alice, 100);
    // Both deletes became overlay records (dropped from core); only the merge runs.
    const adds = r.reviewOps.filter((o) => o.type === "addSuggestion");
    expect(adds.length).toBe(3); // two delete records + one structural (merge)
    const kinds = adds.map((o) => (o.type === "addSuggestion" ? o.s.kind : "")).sort();
    expect(kinds).toEqual(["delete", "delete", "structural"]);

    const { doc, review } = commitInterceptResult(original, r);
    // Merge ran (one paragraph) but no text was destroyed.
    expect(bodyText(doc)).toEqual(["helloworld"]);

    const res = rejectAllSuggestions(doc, review);
    let d = doc;
    for (const op of res.ops) d = applyOp(d, op).doc;
    expect(bodyText(d)).toEqual(["hello", "world"]); // structure + text restored
  });
});
