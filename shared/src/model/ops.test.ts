import { describe, expect, it } from "vitest";
import type { CharStyle, Document, FieldDef, Paragraph, SectionProps } from "./document";
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

describe("styleEq — inline field marker", () => {
  it("runs differing only in fieldId are NOT equal (boundaries survive normalization)", () => {
    expect(styleEq({ ...CHAR, fieldId: "f0" }, { ...CHAR })).toBe(false);
    expect(styleEq({ ...CHAR, fieldId: "f0" }, { ...CHAR, fieldId: "f1" })).toBe(false);
    expect(styleEq({ ...CHAR, fieldId: "f0" }, { ...CHAR, fieldId: "f0" })).toBe(true);
  });
});
