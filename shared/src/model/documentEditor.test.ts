import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, Run, SdtProps, SectionProps } from "./document";
import { DocumentEditor } from "./documentEditor";
import { getParagraphs, sdtText, textOf } from "./query";

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

// --- SDT (content control) editing — the primary templating surface ----------

const runOf = (text: string, sdtPath?: string[]): Run => ({ text, style: sdtPath ? { ...CHAR, sdtPath } : CHAR });

const mkPara = (id: string, runs: Run[], extra: Partial<Paragraph> = {}): Paragraph => ({
  kind: "paragraph",
  id,
  revision: 0,
  runs,
  style: { ...PARA },
  ...extra,
});

describe("DocumentEditor — SDT props", () => {
  const docWith = (props: SdtProps): Document => ({
    section: section(),
    blocks: [mkPara("p", [runOf("before "), runOf("val", ["c"]), runOf(" after")])],
    sdts: { c: props },
  });

  it("setSdtProps merges the patch and preserves type; undo restores", () => {
    const ed = new DocumentEditor(docWith({ type: "plainText", tag: "T" }));
    ed.setSdtProps("c", { alias: "Full Name", tag: "Name" });
    expect(ed.doc.sdts!.c).toEqual({ type: "plainText", tag: "Name", alias: "Full Name" });
    ed.undo();
    expect(ed.doc.sdts!.c).toEqual({ type: "plainText", tag: "T" });
  });

  it("setSdtProps never changes the control type, even for an untyped caller", () => {
    const ed = new DocumentEditor(docWith({ type: "plainText", tag: "T" }));
    // `type` is excluded from the patch shape; force it through as an untyped
    // runtime caller would, and confirm it is ignored.
    ed.setSdtProps("c", { type: "checkbox", tag: "N" } as unknown as Partial<Omit<SdtProps, "type">>);
    expect(ed.doc.sdts!.c!.type).toBe("plainText");
    expect(ed.doc.sdts!.c!.tag).toBe("N");
  });

  it("setCheckbox flips a checkbox control's state", () => {
    const ed = new DocumentEditor(docWith({ type: "checkbox", checked: false }));
    ed.setCheckbox("c", true);
    expect(ed.doc.sdts!.c!.checked).toBe(true);
  });

  it("setCheckbox throws for a non-checkbox control", () => {
    const ed = new DocumentEditor(docWith({ type: "plainText" }));
    expect(() => ed.setCheckbox("c", true)).toThrow(/not a checkbox/);
  });

  it("throws on unknown control ids", () => {
    const ed = new DocumentEditor(docWith({ type: "plainText" }));
    expect(() => ed.setSdtProps("nope", {})).toThrow(/not found/);
    expect(() => ed.setSdtText("nope", "x")).toThrow(/not found/);
  });
});

describe("DocumentEditor — setSdtText", () => {
  it("fills an inline control, keeping surrounding text and the control tag", () => {
    const doc: Document = {
      section: section(),
      blocks: [mkPara("p", [runOf("Name: "), runOf("Ada", ["c"]), runOf("!")])],
      sdts: { c: { type: "plainText", tag: "Name" } },
    };
    const ed = new DocumentEditor(doc);
    ed.setSdtText("c", "Bob");
    const p = ed.getParagraph("p")!;
    expect(textOf(p)).toBe("Name: Bob!");
    expect(sdtText(ed.doc, "c")).toBe("Bob");
    // the replacement run is still tagged with the control id
    expect(p.runs.find((r) => r.text === "Bob")!.style.sdtPath).toEqual(["c"]);
    ed.undo();
    expect(sdtText(ed.doc, "c")).toBe("Ada");
  });

  it("preserves ancestry when filling a NESTED inline control", () => {
    const doc: Document = {
      section: section(),
      blocks: [mkPara("p", [runOf("x"), runOf("42", ["outer", "inner"]), runOf("y")])],
      sdts: { outer: { type: "richText" }, inner: { type: "plainText" } },
    };
    const ed = new DocumentEditor(doc);
    ed.setSdtText("inner", "99");
    const run = ed.getParagraph("p")!.runs.find((r) => r.text === "99")!;
    expect(run.style.sdtPath).toEqual(["outer", "inner"]); // outer control survives
    expect(sdtText(ed.doc, "inner")).toBe("99");
    expect(sdtText(ed.doc, "outer")).toBe("99");
  });

  it("fills a block-level control over a single paragraph, keeping its sdtPath", () => {
    const doc: Document = {
      section: section(),
      blocks: [mkPara("b", [runOf("old value")], { sdtPath: ["sec"] })],
      sdts: { sec: { type: "richText", tag: "Section" } },
    };
    const ed = new DocumentEditor(doc);
    ed.setSdtText("sec", "new value");
    const b = ed.getParagraph("b")!;
    expect(textOf(b)).toBe("new value");
    expect(b.sdtPath).toEqual(["sec"]); // block-level membership intact
  });

  it("clears the placeholder flag when filling a control", () => {
    const doc: Document = {
      section: section(),
      blocks: [mkPara("p", [runOf("placeholder", ["c"])])],
      sdts: { c: { type: "plainText", placeholder: true } },
    };
    const ed = new DocumentEditor(doc);
    ed.setSdtText("c", "real");
    expect(ed.doc.sdts!.c!.placeholder).toBe(false);
    ed.undo();
    expect(ed.doc.sdts!.c!.placeholder).toBe(true); // both ops undo as one step
  });

  it("throws for a control spanning multiple blocks", () => {
    const doc: Document = {
      section: section(),
      blocks: [mkPara("b0", [runOf("one")], { sdtPath: ["sec"] }), mkPara("b1", [runOf("two")], { sdtPath: ["sec"] })],
      sdts: { sec: { type: "richText" } },
    };
    const ed = new DocumentEditor(doc);
    expect(() => ed.setSdtText("sec", "x")).toThrow(/spans 2 block/);
  });
});

describe("DocumentEditor — setSdtValue", () => {
  const listDoc = (type: "dropDown" | "comboBox"): Document => ({
    section: section(),
    blocks: [mkPara("p", [runOf("One", ["c"])])],
    sdts: { c: { type, listItems: [{ display: "One", value: "1" }, { display: "Two", value: "2" }] } },
  });

  it("selects a dropdown option by value (text becomes the display)", () => {
    const ed = new DocumentEditor(listDoc("dropDown"));
    ed.setSdtValue("c", "2");
    expect(sdtText(ed.doc, "c")).toBe("Two");
  });

  it("selects by display when no value matches", () => {
    const ed = new DocumentEditor(listDoc("dropDown"));
    ed.setSdtValue("c", "Two");
    expect(sdtText(ed.doc, "c")).toBe("Two");
  });

  it("throws for a value that is not a dropdown option", () => {
    const ed = new DocumentEditor(listDoc("dropDown"));
    expect(() => ed.setSdtValue("c", "nope")).toThrow(/not an option/);
  });

  it("allows free text on a combo box", () => {
    const ed = new DocumentEditor(listDoc("comboBox"));
    ed.setSdtValue("c", "Freeform");
    expect(sdtText(ed.doc, "c")).toBe("Freeform");
  });

  it("throws for a non-list control", () => {
    const ed = new DocumentEditor({ section: section(), blocks: [mkPara("p", [runOf("x", ["c"])])], sdts: { c: { type: "plainText" } } });
    expect(() => ed.setSdtValue("c", "x")).toThrow(/use setSdtText/);
  });
});

describe("DocumentEditor — removeSdt", () => {
  it("unwraps an inline control: strips the tag, keeps text, deletes props; undo restores", () => {
    const doc: Document = {
      section: section(),
      blocks: [mkPara("p", [runOf("A "), runOf("val", ["c"]), runOf(" B")])],
      sdts: { c: { type: "plainText", tag: "T" } },
    };
    const ed = new DocumentEditor(doc);
    ed.removeSdt("c");
    expect(textOf(ed.getParagraph("p")!)).toBe("A val B"); // content preserved
    expect(ed.getParagraph("p")!.runs.every((r) => r.style.sdtPath === undefined)).toBe(true);
    expect(ed.doc.sdts?.c).toBeUndefined(); // props deleted (invariant kept)
    ed.undo();
    expect(ed.doc.sdts!.c).toEqual({ type: "plainText", tag: "T" });
    expect(sdtText(ed.doc, "c")).toBe("val");
  });

  it("unwrapping an OUTER control keeps the nested inner control", () => {
    const doc: Document = {
      section: section(),
      blocks: [mkPara("p", [runOf("42", ["outer", "inner"])])],
      sdts: { outer: { type: "richText" }, inner: { type: "plainText" } },
    };
    const ed = new DocumentEditor(doc);
    ed.removeSdt("outer");
    expect(ed.doc.sdts?.outer).toBeUndefined();
    expect(ed.doc.sdts!.inner).toBeDefined();
    // inner survives with its path shortened to just [inner]
    expect(ed.getParagraph("p")!.runs[0]!.style.sdtPath).toEqual(["inner"]);
    expect(sdtText(ed.doc, "inner")).toBe("42");
  });

  it("deleteContents removes the tagged runs too", () => {
    const doc: Document = {
      section: section(),
      blocks: [mkPara("p", [runOf("keep "), runOf("gone", ["c"])])],
      sdts: { c: { type: "plainText" } },
    };
    const ed = new DocumentEditor(doc);
    ed.removeSdt("c", { deleteContents: true });
    expect(textOf(ed.getParagraph("p")!)).toBe("keep ");
    expect(ed.doc.sdts?.c).toBeUndefined();
  });

  it("unwraps a block-level control on a top-level body paragraph", () => {
    const doc: Document = {
      section: section(),
      blocks: [mkPara("b", [runOf("body value")], { sdtPath: ["sec"] })],
      sdts: { sec: { type: "richText" } },
    };
    const ed = new DocumentEditor(doc);
    ed.removeSdt("sec");
    const b = ed.getParagraph("b")!;
    expect(b.sdtPath).toBeUndefined();
    expect(textOf(b)).toBe("body value");
    expect(ed.doc.sdts?.sec).toBeUndefined();
  });

  it("deleteContents removes a block-level member block", () => {
    const doc: Document = {
      section: section(),
      blocks: [mkPara("keep", [runOf("stay")]), mkPara("b", [runOf("go")], { sdtPath: ["sec"] })],
      sdts: { sec: { type: "richText" } },
    };
    const ed = new DocumentEditor(doc);
    ed.removeSdt("sec", { deleteContents: true });
    expect(ed.doc.blocks.map((x) => x.id)).toEqual(["keep"]);
    expect(ed.doc.sdts?.sec).toBeUndefined();
  });

  it("throws for a block-level member inside a table cell", () => {
    const doc: Document = {
      section: section(),
      blocks: [
        { kind: "table", id: "t", revision: 0, rows: [{ cells: [{ id: "c0", blocks: [mkPara("cp", [runOf("x")], { sdtPath: ["sec"] })] }] }] },
      ],
      sdts: { sec: { type: "richText" } },
    };
    const ed = new DocumentEditor(doc);
    expect(() => ed.removeSdt("sec")).toThrow(/table cell/);
  });

  it("throws on unknown control ids", () => {
    const ed = new DocumentEditor({ section: section(), blocks: [mkPara("p", [runOf("x")])] });
    expect(() => ed.removeSdt("nope")).toThrow(/not found/);
  });
});
