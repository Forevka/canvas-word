// Nested content controls (w:sdt inside w:sdt) must import to an ORDERED sdtPath
// ancestry: inline controls on the runs (CharStyle.sdtPath), block-level controls
// on the blocks (Block.sdtPath). A run's full chain is block.sdtPath ++ run path.

import { describe, expect, it } from "vitest";
import type { Document, Paragraph } from "@cw/shared";
import { runImport } from "./pipeline";
import { simpleDocx } from "./fixture";

const txt = (s: string): string => `<w:r><w:t>${s}</w:t></w:r>`;
const sdt = (tag: string, inner: string, extraPr = ""): string =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/>${extraPr}</w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const paras = (blocks: { kind: string }[]): Paragraph[] =>
  blocks.filter((b): b is Paragraph => b.kind === "paragraph");

/** Map an id path to the tags of those controls (stable identity for asserting). */
const tagsOf = (doc: Document, path: string[] | undefined): string[] =>
  (path ?? []).map((id) => doc.sdts![id]!.tag!);
const runChainTags = (doc: Document, block: Paragraph, text: string): string[] => {
  const r = block.runs.find((rn) => rn.text === text)!;
  return [...tagsOf(doc, block.sdtPath), ...tagsOf(doc, r.style.sdtPath)];
};

describe("nested SDT import", () => {
  it("inline-in-inline → run carries [outer, inner]", () => {
    const body = `<w:p>${sdt("outer", txt("a") + sdt("inner", txt("b")) + txt("c"))}</w:p>`;
    const doc = runImport(simpleDocx(body)).doc;
    const p = paras(doc.blocks)[0]!;
    expect(runChainTags(doc, p, "a")).toEqual(["outer"]);
    expect(runChainTags(doc, p, "b")).toEqual(["outer", "inner"]);
    expect(runChainTags(doc, p, "c")).toEqual(["outer"]);
  });

  it("block-in-block → blocks carry [outer] / [outer, inner]", () => {
    const body = sdt("sec", `<w:p>${txt("one")}</w:p>` + sdt("sub", `<w:p>${txt("two")}</w:p>`) + `<w:p>${txt("three")}</w:p>`);
    const doc = runImport(simpleDocx(body)).doc;
    const ps = paras(doc.blocks);
    expect(tagsOf(doc, ps.find((p) => p.runs.some((r) => r.text === "one"))!.sdtPath)).toEqual(["sec"]);
    expect(tagsOf(doc, ps.find((p) => p.runs.some((r) => r.text === "two"))!.sdtPath)).toEqual(["sec", "sub"]);
    expect(tagsOf(doc, ps.find((p) => p.runs.some((r) => r.text === "three"))!.sdtPath)).toEqual(["sec"]);
  });

  it("inline-in-block → block path on the paragraph, inline path on the run (split levels)", () => {
    const body = sdt("sec", `<w:p>${txt("x")}${sdt("f", txt("y"))}</w:p>`);
    const doc = runImport(simpleDocx(body)).doc;
    const p = paras(doc.blocks)[0]!;
    expect(tagsOf(doc, p.sdtPath)).toEqual(["sec"]); // block-level
    expect(tagsOf(doc, p.runs.find((r) => r.text === "x")!.style.sdtPath)).toEqual([]); // not in the inline control
    expect(tagsOf(doc, p.runs.find((r) => r.text === "y")!.style.sdtPath)).toEqual(["f"]); // inline only
    expect(runChainTags(doc, p, "y")).toEqual(["sec", "f"]); // full chain
  });

  it("section wrapping a TABLE → table carries the section; cell paragraphs do NOT re-inherit it", () => {
    const cell = (inner: string): string => `<w:tc><w:tcPr/><w:p>${inner}</w:p></w:tc>`;
    const table = `<w:tbl><w:tblPr/><w:tblGrid/><w:tr>${cell(txt("Label"))}${cell(sdt("val", txt("200")))}</w:tr></w:tbl>`;
    const body = sdt("sec", table);
    const doc = runImport(simpleDocx(body)).doc;
    const tbl = doc.blocks.find((b) => b.kind === "table")!;
    expect(tagsOf(doc, tbl.sdtPath)).toEqual(["sec"]); // section wraps the table
    const cellPara = tbl.kind === "table" ? (tbl.rows[0]!.cells[1]!.blocks[0] as Paragraph) : undefined!;
    expect(cellPara.sdtPath).toBeUndefined(); // cell does NOT re-carry the section
    expect(tagsOf(doc, cellPara.runs.find((r) => r.text === "200")!.style.sdtPath)).toEqual(["val"]); // inner inline control
  });

  it("checkbox as the inner (leaf) control normalizes the glyph", () => {
    const checkboxPr = `<w14:checkbox><w14:checked w14:val="1"/></w14:checkbox>`;
    const body = `<w:p>${sdt("outer", sdt("cb", txt("X"), checkboxPr))}</w:p>`;
    const doc = runImport(simpleDocx(body)).doc;
    const p = paras(doc.blocks)[0]!;
    const r = p.runs.find((rn) => rn.text === "☒")!; // glyph-normalized from checkbox state
    expect(tagsOf(doc, r.style.sdtPath)).toEqual(["outer", "cb"]);
  });
});
