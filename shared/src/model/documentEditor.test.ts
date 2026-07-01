import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, SectionProps } from "./document";
import { DocumentEditor } from "./documentEditor";
import { getParagraphs, textOf } from "./query";

const CHAR: CharStyle = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA = { align: "left" as const, lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };

const para = (id: string, text: string, extra: Partial<Paragraph["style"]> = {}): Paragraph => ({
  kind: "paragraph",
  id,
  revision: 0,
  runs: text.length > 0 ? [{ text, style: CHAR }] : [],
  style: { ...PARA, ...extra },
});

const section = (): SectionProps => ({ pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } });
const docOf = (...ps: Paragraph[]): Document => ({ section: section(), blocks: ps });

const textAt = (ed: DocumentEditor, id: string) => textOf(ed.getParagraph(id)!);

describe("DocumentEditor — text editing", () => {
  it("setParagraphText replaces run text and does not mutate the input", () => {
    const input = docOf(para("p", "hello"));
    const ed = new DocumentEditor(input);
    ed.setParagraphText("p", "world");
    expect(textAt(ed, "p")).toBe("world");
    expect(input.blocks[0]).toMatchObject({ runs: [{ text: "hello" }] }); // input untouched
    expect(ed.doc).not.toBe(input);
  });

  it("insertText / deleteText / replaceText operate on UTF-16 offsets", () => {
    const ed = new DocumentEditor(docOf(para("p", "hello world")));
    ed.insertText("p", 5, ",");
    expect(textAt(ed, "p")).toBe("hello, world");
    ed.deleteText("p", 0, 7);
    expect(textAt(ed, "p")).toBe("world");
    ed.replaceText("p", 0, 5, "earth");
    expect(textAt(ed, "p")).toBe("earth");
  });

  it("setParagraphStyle patches the paragraph style", () => {
    const ed = new DocumentEditor(docOf(para("p", "x")));
    ed.setParagraphStyle("p", { align: "center" });
    expect(ed.getParagraph("p")!.style.align).toBe("center");
  });

  it("throws on unknown paragraph ids", () => {
    const ed = new DocumentEditor(docOf(para("p", "x")));
    expect(() => ed.setParagraphText("nope", "y")).toThrow(/not found/);
  });

  it("clears an already-empty paragraph without requiring a style", () => {
    const ed = new DocumentEditor(docOf(para("p", "")));
    expect(() => ed.setParagraphText("p", "")).not.toThrow();
    expect(ed.getParagraph("p")!.runs).toEqual([]);
  });
});

describe("DocumentEditor — structural editing", () => {
  it("insertParagraph after clones the reference style (minus section break) and mints an id", () => {
    const ed = new DocumentEditor(docOf(para("a", "first", { align: "right", sectionBreak: { type: "nextPage", props: {} } }), para("b", "second")));
    ed.insertParagraph("a", "inserted");
    const id = ed.lastInsertedId!;
    const ids = ed.doc.blocks.map((x) => x.id);
    expect(ids).toEqual(["a", id, "b"]);
    const inserted = ed.getParagraph(id)!;
    expect(textOf(inserted)).toBe("inserted");
    expect(inserted.style.align).toBe("right"); // inherited
    expect(inserted.style.sectionBreak).toBeUndefined(); // structural marker stripped
  });

  it("insertParagraph before inserts at the reference index", () => {
    const ed = new DocumentEditor(docOf(para("a", "first"), para("b", "second")));
    ed.insertParagraph("b", "middle", { position: "before" });
    expect(ed.doc.blocks.map((x) => x.id)).toEqual(["a", ed.lastInsertedId, "b"]);
  });

  it("removeBlock drops a top-level block", () => {
    const ed = new DocumentEditor(docOf(para("a", "first"), para("b", "second")));
    ed.removeBlock("a");
    expect(ed.doc.blocks.map((x) => x.id)).toEqual(["b"]);
  });

  it("insertParagraph throws when the reference is not a top-level block", () => {
    const doc: Document = {
      section: section(),
      blocks: [
        para("p", "body"),
        { kind: "table", id: "t", revision: 0, rows: [{ cells: [{ id: "c0", blocks: [para("cellPara", "in cell")] }] }] },
      ],
    };
    const ed = new DocumentEditor(doc);
    expect(() => ed.insertParagraph("cellPara", "x")).toThrow(/not a top-level block/);
  });

  it("insertParagraph throws when the reference is not a paragraph and no style is given", () => {
    const doc: Document = {
      section: section(),
      blocks: [{ kind: "table", id: "t", revision: 0, rows: [{ cells: [{ id: "c0", blocks: [para("cellPara", "in cell")] }] }] }],
    };
    const ed = new DocumentEditor(doc);
    expect(() => ed.insertParagraph("t", "x")).toThrow(/not a paragraph/);
  });
});

describe("DocumentEditor — undo / redo", () => {
  it("undo reverts the last edit; redo re-applies it", () => {
    const ed = new DocumentEditor(docOf(para("p", "hello")));
    ed.setParagraphText("p", "world");
    expect(ed.canUndo).toBe(true);
    expect(ed.undo()).toBe(true);
    expect(textAt(ed, "p")).toBe("hello");
    expect(ed.canRedo).toBe(true);
    expect(ed.redo()).toBe(true);
    expect(textAt(ed, "p")).toBe("world");
  });

  it("a multi-op edit (replaceText) undoes as a single step", () => {
    const ed = new DocumentEditor(docOf(para("p", "hello world")));
    ed.replaceText("p", 0, 5, "goodbye");
    expect(textAt(ed, "p")).toBe("goodbye world");
    ed.undo();
    expect(textAt(ed, "p")).toBe("hello world");
  });

  it("returns false when there is nothing to undo/redo", () => {
    const ed = new DocumentEditor(docOf(para("p", "x")));
    expect(ed.undo()).toBe(false);
    expect(ed.redo()).toBe(false);
  });

  it("a new edit clears the redo stack", () => {
    const ed = new DocumentEditor(docOf(para("p", "hello")));
    ed.setParagraphText("p", "world");
    ed.undo();
    ed.setParagraphText("p", "earth");
    expect(ed.canRedo).toBe(false);
    expect(textAt(ed, "p")).toBe("earth");
  });

  it("insertParagraph then undo removes the inserted block", () => {
    const ed = new DocumentEditor(docOf(para("a", "first")));
    ed.insertParagraph("a", "second");
    expect(getParagraphs(ed.doc)).toHaveLength(2);
    ed.undo();
    expect(getParagraphs(ed.doc).map((p) => p.id)).toEqual(["a"]);
  });
});

describe("DocumentEditor — chaining & find", () => {
  it("edits chain and find locates matches on the live doc", () => {
    const ed = new DocumentEditor(docOf(para("a", "alpha"), para("b", "beta")));
    ed.setParagraphText("a", "gamma").setParagraphText("b", "gamma two");
    expect(ed.find("gamma").map((m) => m.paragraph.id)).toEqual(["a", "b"]);
  });
});
