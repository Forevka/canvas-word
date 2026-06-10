import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, SectionProps } from "@cw/shared";
import { applyOp } from "@cw/shared";
import { addBookmarkCmd, removeBookmarkCmd, renameBookmarkCmd } from "./commands";
import type { EditorState } from "./state";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 14, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA: ParaStyle = { align: "left", lineHeight: 1.35, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

let n = 0;
const p = (text: string): Paragraph => ({ kind: "paragraph", id: `p${n++}`, revision: 0, runs: [{ text, style: CHAR }], style: PARA });
const docOf = (...ps: Paragraph[]): Document => ({ section: SECTION, blocks: ps });

const apply = (state: EditorState, cmd: ReturnType<typeof addBookmarkCmd>): Document => {
  const trn = cmd(state);
  if (!trn) throw new Error("no transaction");
  let doc = state.doc;
  for (const op of trn.ops) doc = applyOp(doc, op).doc;
  return doc;
};

describe("bookmark commands", () => {
  it("adds a bookmark covering the selection", () => {
    const a = p("hello world");
    const state: EditorState = { doc: docOf(a), selection: { anchor: { blockId: a.id, offset: 0 }, focus: { blockId: a.id, offset: 5 } } };
    const doc = apply(state, addBookmarkCmd("bm"));
    expect(doc.bookmarks!["bm"]).toEqual({ start: { blockId: a.id, offset: 0 }, end: { blockId: a.id, offset: 5 } });
  });

  it("renames a bookmark, preserving its range, and removes it", () => {
    const a = p("hello");
    let doc = docOf(a);
    doc = { ...doc, bookmarks: { old: { start: { blockId: a.id, offset: 1 }, end: { blockId: a.id, offset: 4 } } } };
    const renamed = apply({ doc, selection: null }, renameBookmarkCmd("old", "fresh"));
    expect(renamed.bookmarks!["old"]).toBeUndefined();
    expect(renamed.bookmarks!["fresh"]).toEqual({ start: { blockId: a.id, offset: 1 }, end: { blockId: a.id, offset: 4 } });
    const removed = apply({ doc: renamed, selection: null }, removeBookmarkCmd("fresh"));
    expect(removed.bookmarks!["fresh"]).toBeUndefined();
  });

  it("a setBookmark op carries its exact inverse (undo)", () => {
    const a = p("x");
    const base = docOf(a);
    const range = { start: { blockId: a.id, offset: 0 }, end: { blockId: a.id, offset: 1 } };
    const res = applyOp(base, { type: "setBookmark", name: "b", range });
    expect(res.doc.bookmarks!["b"]).toEqual(range);
    const undone = applyOp(res.doc, res.inverse).doc;
    expect(undone.bookmarks?.["b"]).toBeUndefined();
  });

  it("rebases a bookmark offset past text inserted before it (via mapPosition)", () => {
    const a = p("hello");
    const start = { blockId: a.id, offset: 3 };
    const res = applyOp(docOf(a), { type: "insertText", at: { blockId: a.id, offset: 0 }, text: "ABC" });
    expect(res.mapPosition(start)).toEqual({ blockId: a.id, offset: 6 }); // 3 + len("ABC")
  });
});
