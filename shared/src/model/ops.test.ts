import { describe, expect, it } from "vitest";
import type { CharStyle, Document, FieldDef, Paragraph, SectionProps, TableBlock, TableCell, TableRow } from "./document";
import { applyOp, styleEq } from "./ops";

const section = (): SectionProps => ({ pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } });
const CHAR: CharStyle = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA = { align: "left" as const, lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const para = (id: string, text: string): Paragraph => ({ kind: "paragraph", id, revision: 0, runs: [{ text, style: CHAR }], style: PARA });
const docOf = (...ps: Paragraph[]): Document => ({ section: section(), blocks: ps });

const pageField = (id: string): FieldDef => ({ id, instruction: " PAGE \\* roman ", name: "PAGE", kind: "builtin", spec: { type: "PAGE", numFmt: "roman" } });

describe("applyOp setField", () => {
  it("registers a field def; its inverse removes it", () => {
    const def = pageField("f0");
    const r = applyOp(docOf(para("p", "x")), { type: "setField", id: "f0", def });
    expect(r.doc.fields?.["f0"]).toEqual(def);
    expect(applyOp(r.doc, r.inverse).doc.fields?.["f0"]).toBeUndefined();
  });

  it("removing a field def (def:null) round-trips via inverse", () => {
    const def = pageField("f0");
    const d: Document = { ...docOf(para("p", "x")), fields: { f0: def } };
    const r = applyOp(d, { type: "setField", id: "f0", def: null });
    expect(r.doc.fields?.["f0"]).toBeUndefined();
    expect(applyOp(r.doc, r.inverse).doc.fields?.["f0"]).toEqual(def);
  });

  it("does not mutate the input document", () => {
    const d = docOf(para("p", "x"));
    applyOp(d, { type: "setField", id: "f0", def: pageField("f0") });
    expect(d.fields).toBeUndefined();
  });
});

describe("applyOp setTocInstruction", () => {
  it("sets the doc TOC instruction; inverse restores the prior (undefined) value", () => {
    const r = applyOp(docOf(para("p", "x")), { type: "setTocInstruction", instruction: ' TOC \\o "1-5" \\h ' });
    expect(r.doc.tocInstruction).toBe(' TOC \\o "1-5" \\h ');
    expect(applyOp(r.doc, r.inverse).doc.tocInstruction).toBeUndefined();
  });

  it("replacing an existing instruction round-trips via inverse", () => {
    const d: Document = { ...docOf(para("p", "x")), tocInstruction: ' TOC \\o "1-3" ' };
    const r = applyOp(d, { type: "setTocInstruction", instruction: ' TOC \\o "1-2" \\z ' });
    expect(r.doc.tocInstruction).toBe(' TOC \\o "1-2" \\z ');
    expect(applyOp(r.doc, r.inverse).doc.tocInstruction).toBe(' TOC \\o "1-3" ');
  });

  it("does not mutate the input document", () => {
    const d = docOf(para("p", "x"));
    applyOp(d, { type: "setTocInstruction", instruction: " TOC " });
    expect(d.tocInstruction).toBeUndefined();
  });
});

describe("applyOp splitParagraph — block-level content control", () => {
  it("carries newSdtPath onto the tail so Enter stays inside a block-level SDT", () => {
    const src: Paragraph = { ...para("p", "hello world"), sdtPath: ["sec"] };
    const r = applyOp(docOf(src), {
      type: "splitParagraph",
      at: { blockId: "p", offset: 5 },
      newBlockId: "p2",
      newSdtPath: ["sec"],
    });
    const [head, tail] = r.doc.blocks as Paragraph[];
    expect(head!.sdtPath).toEqual(["sec"]); // head keeps it (untouched)
    expect(tail!.sdtPath).toEqual(["sec"]); // tail inherits it
  });

  it("a split with no newSdtPath leaves the tail outside any block control", () => {
    const r = applyOp(docOf(para("p", "hello world")), {
      type: "splitParagraph",
      at: { blockId: "p", offset: 5 },
      newBlockId: "p2",
    });
    expect((r.doc.blocks[1] as Paragraph).sdtPath).toBeUndefined();
  });

  it("merge inverse restores the tail's block-level control (undo round-trip)", () => {
    const head: Paragraph = { ...para("p", "hello "), sdtPath: ["sec"] };
    const tail: Paragraph = { ...para("p2", "world"), sdtPath: ["sec"] };
    const merge = applyOp(docOf(head, tail), { type: "mergeParagraphs", firstBlockId: "p" });
    expect(merge.doc.blocks).toHaveLength(1);
    // Undo: the re-split must put the control back on the restored tail.
    const undo = applyOp(merge.doc, merge.inverse);
    expect((undo.doc.blocks[1] as Paragraph).sdtPath).toEqual(["sec"]);
  });
});

describe("applyOp setRowHeight — interactive row-drag height (w:trHeight)", () => {
  const cell = (id: string, text: string): TableCell => ({ id, blocks: [para(id, text)] });
  const tableDoc = (...rows: TableRow[]): Document => ({
    section: section(),
    blocks: [{ kind: "table", id: "t0", revision: 0, rows }],
  });
  const firstRow = (d: Document): TableRow => (d.blocks[0] as TableBlock).rows[0]!;

  it("writes a row's height as { value, rule:'atLeast' }; inverse clears it (was none)", () => {
    const d = tableDoc({ cells: [cell("a", "x")] }, { cells: [cell("b", "y")] });
    const r = applyOp(d, { type: "setRowHeight", tableId: "t0", rowIndex: 0, height: { value: 80, rule: "atLeast" } });
    expect(firstRow(r.doc).props?.height).toEqual({ value: 80, rule: "atLeast" });
    // the second row is untouched
    expect((r.doc.blocks[0] as TableBlock).rows[1]!.props).toBeUndefined();
    const undo = applyOp(r.doc, r.inverse);
    expect(firstRow(undo.doc).props?.height).toBeUndefined();
  });

  it("replacing an existing height round-trips via inverse, preserving sibling props", () => {
    const d = tableDoc({ cells: [cell("a", "x")], props: { height: { value: 50, rule: "exact" }, cantSplit: true } });
    const r = applyOp(d, { type: "setRowHeight", tableId: "t0", rowIndex: 0, height: { value: 120, rule: "atLeast" } });
    expect(firstRow(r.doc).props).toEqual({ height: { value: 120, rule: "atLeast" }, cantSplit: true });
    const undo = applyOp(r.doc, r.inverse);
    expect(firstRow(undo.doc).props).toEqual({ height: { value: 50, rule: "exact" }, cantSplit: true });
  });

  it("clearing height (null) drops the props object when nothing else remains", () => {
    const d = tableDoc({ cells: [cell("a", "x")], props: { height: { value: 60, rule: "atLeast" } } });
    const r = applyOp(d, { type: "setRowHeight", tableId: "t0", rowIndex: 0, height: null });
    expect(firstRow(r.doc).props).toBeUndefined();
    // inverse restores the prior height
    expect(firstRow(applyOp(r.doc, r.inverse).doc).props?.height).toEqual({ value: 60, rule: "atLeast" });
  });
});

describe("styleEq — inline field marker", () => {
  it("runs differing only in fieldId are NOT equal (boundaries survive normalization)", () => {
    expect(styleEq({ ...CHAR, fieldId: "f0" }, { ...CHAR })).toBe(false);
    expect(styleEq({ ...CHAR, fieldId: "f0" }, { ...CHAR, fieldId: "f1" })).toBe(false);
    expect(styleEq({ ...CHAR, fieldId: "f0" }, { ...CHAR, fieldId: "f0" })).toBe(true);
  });
});
