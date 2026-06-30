import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Op, Paragraph, ParaStyle, Run, SectionProps } from "@cw/shared";
import { applyOp } from "@cw/shared";
import { insertEndnoteCmd } from "./commands";
import type { EditorState } from "./state";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 14, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA: ParaStyle = { align: "left", lineHeight: 1.35, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

let n = 0;
const p = (...runs: Run[]): Paragraph => ({ kind: "paragraph", id: `p${n++}`, revision: 0, runs, style: PARA });
const run = (text: string, style: Partial<CharStyle> = {}): Run => ({ text, style: { ...CHAR, ...style } });
const docOf = (...ps: Paragraph[]): Document => ({ section: SECTION, blocks: ps });

/** Apply a command, returning the resulting doc plus the inverse ops (for undo). */
const run_ = (state: EditorState, cmd: ReturnType<typeof insertEndnoteCmd>): { doc: Document; inverses: Op[] } => {
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

const refRun = (doc: Document): Run | undefined => doc.blocks.flatMap((b) => (b.kind === "paragraph" ? b.runs : [])).find((r) => r.style.endnoteRef);

describe("insertEndnoteCmd", () => {
  it("inserts a superscript reference at the caret and creates the endnote body", () => {
    const a = p(run("hello"));
    const state: EditorState = { doc: docOf(a), selection: { anchor: { blockId: a.id, offset: 5 }, focus: { blockId: a.id, offset: 5 } } };
    const { doc } = run_(state, insertEndnoteCmd());

    const ref = refRun(doc);
    expect(ref).toBeDefined();
    expect(ref!.text).toBe("1");
    expect(ref!.style.verticalAlign).toBe("super");
    const noteId = ref!.style.endnoteRef!;
    expect(doc.endnotes?.[noteId]).toBeDefined();
    expect(doc.endnotes![noteId]).toHaveLength(1);
    // The body is empty and ready for typing; footnotes are untouched.
    expect(doc.endnotes![noteId]![0]!.runs.map((r) => r.text).join("")).toBe("");
    expect(doc.footnotes).toBeUndefined();
  });

  it("is a no-op without a collapsed body selection", () => {
    const a = p(run("x"));
    expect(insertEndnoteCmd()({ doc: docOf(a), selection: null })).toBeNull();
    // Non-collapsed range → no insert.
    expect(insertEndnoteCmd()({ doc: docOf(a), selection: { anchor: { blockId: a.id, offset: 0 }, focus: { blockId: a.id, offset: 1 } } })).toBeNull();
  });

  it("renumbers a later endnote reference when one is inserted before it", () => {
    // Paragraph already carries an endnote ref "1" after the word "tail".
    const a = p(run("head "), run("tail"), run("1", { verticalAlign: "super", endnoteRef: "en_existing" }));
    const before = { ...docOf(a), endnotes: { en_existing: [p(run("old note", { fontSizePx: 12 }))] } };
    // Caret sits at offset 0 — before the existing ref.
    const state: EditorState = { doc: before, selection: { anchor: { blockId: a.id, offset: 0 }, focus: { blockId: a.id, offset: 0 } } };
    const { doc } = run_(state, insertEndnoteCmd());

    const para = doc.blocks[0] as Paragraph;
    const refs = para.runs.filter((r) => r.style.endnoteRef);
    expect(refs.map((r) => r.text)).toEqual(["1", "2"]); // new ref is 1, the pre-existing one bumps to 2
  });

  it("does not inherit a preceding footnote marker's ref flag", () => {
    // Caret sits right after a footnote marker run, so styleAtRuns echoes its
    // style (footnoteRef + super). The new endnote runs must drop that.
    const a = p(run("body"), run("1", { verticalAlign: "super", footnoteRef: "fn_x" }));
    const state: EditorState = { doc: { ...docOf(a), footnotes: { fn_x: [p(run("a note"))] } }, selection: { anchor: { blockId: a.id, offset: 5 }, focus: { blockId: a.id, offset: 5 } } };
    const { doc } = run_(state, insertEndnoteCmd());

    const ref = refRun(doc);
    expect(ref!.style.footnoteRef).toBeUndefined();
    expect(ref!.style.endnoteRef).toBeDefined();
    const noteRun = doc.endnotes![ref!.style.endnoteRef!]![0]!.runs[0]!;
    expect(noteRun.style.footnoteRef).toBeUndefined();
    expect(noteRun.style.endnoteRef).toBeUndefined();
    expect(noteRun.style.verticalAlign).toBeUndefined();
  });

  it("undoes cleanly, restoring the document to its prior state", () => {
    const a = p(run("hello"));
    const base = docOf(a);
    const state: EditorState = { doc: base, selection: { anchor: { blockId: a.id, offset: 5 }, focus: { blockId: a.id, offset: 5 } } };
    const { doc, inverses } = run_(state, insertEndnoteCmd());
    expect(refRun(doc)).toBeDefined();

    const undone = undo(doc, inverses);
    expect(refRun(undone)).toBeUndefined();
    expect(undone.endnotes ?? {}).toEqual({});
    expect((undone.blocks[0] as Paragraph).runs.map((r) => r.text).join("")).toBe("hello");
  });
});
