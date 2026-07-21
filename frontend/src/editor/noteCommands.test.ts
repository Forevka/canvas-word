import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Op, Paragraph, ParaStyle, Run, SectionProps } from "@cw/shared";
import { applyOp } from "@cw/shared";
import { deleteNoteCmd, noteAtPosition } from "./commands";
import type { EditorState } from "./state";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 14, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA: ParaStyle = { align: "left", lineHeight: 1.35, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

let n = 0;
const p = (...runs: Run[]): Paragraph => ({ kind: "paragraph", id: `p${n++}`, revision: 0, runs, style: PARA });
const run = (text: string, style: Partial<CharStyle> = {}): Run => ({ text, style: { ...CHAR, ...style } });
const docOf = (...ps: Paragraph[]): Document => ({ section: SECTION, blocks: ps });

/** Apply a command, returning the resulting doc plus the inverse ops (for undo). */
const run_ = (state: EditorState, cmd: ReturnType<typeof deleteNoteCmd>): { doc: Document; inverses: Op[] } => {
  const trn = cmd(state);
  if (!trn) throw new Error("no transaction");
  let doc = state.doc;
  const inverses: Op[] = [];
  for (const op of trn.ops) {
    const res = applyOp(doc, op);
    inverses.unshift(res.inverse);
    doc = res.doc;
  }
  return { doc, inverses };
};

const undo = (doc: Document, inverses: Op[]): Document => {
  for (const inv of inverses) doc = applyOp(doc, inv).doc;
  return doc;
};

const refTexts = (doc: Document, key: "footnoteRef" | "endnoteRef"): string[] =>
  doc.blocks.flatMap((b) => (b.kind === "paragraph" ? b.runs : [])).filter((r) => r.style[key]).map((r) => r.text);

describe("noteAtPosition", () => {
  it("detects a footnote marker under the caret and reports its kind + id", () => {
    const a = p(run("body"), run("1", { verticalAlign: "super", footnoteRef: "fn_a" }));
    const doc = { ...docOf(a), footnotes: { fn_a: [p(run("note body"))] } };
    // Caret right after the marker run.
    expect(noteAtPosition(doc, { blockId: a.id, offset: 5 })).toEqual({ kind: "footnote", id: "fn_a" });
    // Caret in the plain text before the marker → no note.
    expect(noteAtPosition(doc, { blockId: a.id, offset: 2 })).toBeNull();
  });

  it("detects an endnote marker and returns null for a non-paragraph / missing block", () => {
    const a = p(run("x"), run("1", { verticalAlign: "super", endnoteRef: "en_a" }));
    const doc = { ...docOf(a), endnotes: { en_a: [p(run("note"))] } };
    expect(noteAtPosition(doc, { blockId: a.id, offset: 2 })).toEqual({ kind: "endnote", id: "en_a" });
    expect(noteAtPosition(doc, { blockId: "nope", offset: 0 })).toBeNull();
  });
});

describe("deleteNoteCmd", () => {
  it("removes the marker run, deletes the body, and renumbers later footnotes", () => {
    // Two footnotes: "1" (fn_a) then "2" (fn_b). Delete the first.
    const a = p(run("alpha"), run("1", { verticalAlign: "super", footnoteRef: "fn_a" }));
    const b = p(run("beta"), run("2", { verticalAlign: "super", footnoteRef: "fn_b" }));
    const doc: Document = { ...docOf(a, b), footnotes: { fn_a: [p(run("first note", { fontSizePx: 12 }))], fn_b: [p(run("second note", { fontSizePx: 12 }))] } };
    const state: EditorState = { doc, selection: { anchor: { blockId: a.id, offset: 6 }, focus: { blockId: a.id, offset: 6 } } };

    const { doc: after } = run_(state, deleteNoteCmd("footnote", "fn_a"));
    // fn_a body gone, fn_b remains and renumbers to "1".
    expect(after.footnotes?.fn_a).toBeUndefined();
    expect(after.footnotes?.fn_b).toBeDefined();
    expect(refTexts(after, "footnoteRef")).toEqual(["1"]);
    // The deleted marker run is gone from paragraph a; its text ("alpha") stays.
    const pa = after.blocks[0] as Paragraph;
    expect(pa.runs.some((r) => r.style.footnoteRef)).toBe(false);
    expect(pa.runs.map((r) => r.text).join("")).toBe("alpha");
  });

  it("keeps an empty run when the marker was the paragraph's only run", () => {
    const a = p(run("1", { verticalAlign: "super", footnoteRef: "fn_only" }));
    const doc: Document = { ...docOf(a), footnotes: { fn_only: [p(run("body"))] } };
    const state: EditorState = { doc, selection: { anchor: { blockId: a.id, offset: 1 }, focus: { blockId: a.id, offset: 1 } } };
    const { doc: after } = run_(state, deleteNoteCmd("footnote", "fn_only"));
    const pa = after.blocks[0] as Paragraph;
    expect(pa.runs).toHaveLength(1);
    expect(pa.runs[0]!.text).toBe("");
    expect(pa.runs[0]!.style.footnoteRef).toBeUndefined();
    expect(pa.runs[0]!.style.verticalAlign).toBeUndefined();
  });

  it("is a no-op for an unknown note id", () => {
    const a = p(run("x"));
    expect(deleteNoteCmd("footnote", "missing")({ doc: docOf(a), selection: null })).toBeNull();
  });

  it("undoes cleanly, restoring marker, numbering and body", () => {
    const a = p(run("one"), run("1", { verticalAlign: "super", endnoteRef: "en_a" }));
    const b = p(run("two"), run("2", { verticalAlign: "super", endnoteRef: "en_b" }));
    const base: Document = { ...docOf(a, b), endnotes: { en_a: [p(run("A"))], en_b: [p(run("B"))] } };
    const state: EditorState = { doc: base, selection: { anchor: { blockId: a.id, offset: 4 }, focus: { blockId: a.id, offset: 4 } } };
    const { doc: after, inverses } = run_(state, deleteNoteCmd("endnote", "en_a"));
    expect(refTexts(after, "endnoteRef")).toEqual(["1"]);

    const undone = undo(after, inverses);
    expect(refTexts(undone, "endnoteRef")).toEqual(["1", "2"]);
    expect(undone.endnotes?.en_a).toBeDefined();
    expect(undone.endnotes?.en_b).toBeDefined();
  });
});
