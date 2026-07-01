import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, SectionProps, TableBlock } from "@cw/shared";
import {
  findText,
  mapPages,
  queryBlockPath,
  queryBookmarks,
  queryEndnotes,
  queryFields,
  queryFootnotes,
  queryListItems,
  queryParagraphs,
  queryPositionOfText,
  queryRangeText,
  querySdts,
  querySdtValue,
  querySections,
  queryStyles,
} from "./queryBridge";
import type { PageInfo as LayoutPageInfo } from "../layout/pages";

const CHAR: CharStyle = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA: ParaStyle = { align: "left", lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };

const para = (id: string, text: string, extra: Partial<ParaStyle> = {}): Paragraph => ({
  kind: "paragraph",
  id,
  revision: 0,
  runs: [{ text, style: CHAR }],
  style: { ...PARA, ...extra },
});

const table = (id: string, ...paras: Paragraph[]): TableBlock => ({
  kind: "table",
  id,
  revision: 0,
  rows: [{ cells: [{ id: `${id}-c0`, blocks: paras }] }],
});

const section = (): SectionProps => ({ pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } });

describe("queryBridge.queryParagraphs", () => {
  it("flattens paragraphs with container/cell context", () => {
    const doc: Document = {
      section: { ...section(), header: [para("h1", "header")] },
      blocks: [para("p1", "body one", { namedStyle: "Heading1", outlineLevel: 0 }), table("t1", para("cellP", "cell text"))],
    };
    const rows = queryParagraphs(doc);
    expect(rows.map((r) => r.id)).toEqual(["p1", "cellP", "h1"]);

    const p1 = rows.find((r) => r.id === "p1")!;
    expect(p1).toMatchObject({ text: "body one", container: "body", tableId: null, row: -1, col: -1, styleName: "Heading1", outlineLevel: 0 });

    const cell = rows.find((r) => r.id === "cellP")!;
    expect(cell).toMatchObject({ container: "body", tableId: "t1", row: 0, col: 0 });

    const header = rows.find((r) => r.id === "h1")!;
    expect(header.container).toBe("header");
    expect(header.styleName).toBeNull();
    expect(header.outlineLevel).toBeNull();
  });
});

describe("queryBridge.findText", () => {
  it("returns only matching paragraphs, flattened", () => {
    const doc: Document = { section: section(), blocks: [para("a", "keep me"), para("b", "drop"), para("c", "keep too")] };
    expect(findText(doc, "keep").map((r) => r.id)).toEqual(["a", "c"]);
  });
});

describe("queryBridge.querySections", () => {
  it("flattens per-section geometry and block ranges", () => {
    const doc: Document = {
      section: section(),
      blocks: [para("a", "one", { sectionBreak: { type: "nextPage", props: {} } }), para("b", "two")],
    };
    const secs = querySections(doc);
    expect(secs).toHaveLength(2);
    expect(secs[0]).toMatchObject({ index: 0, startBlock: 0, endBlock: 0, breakType: "nextPage", pageWidthPx: 816, marginTop: 96, columnCount: 1 });
    expect(secs[1]).toMatchObject({ index: 1, startBlock: 1, endBlock: 1 });
  });
});

describe("queryBridge.querySdts", () => {
  it("flattens the SDT forest with nesting links, props, and text", () => {
    const doc: Document = {
      section: section(),
      blocks: [
        { ...para("b0", "Label: "), runs: [{ text: "Label: ", style: CHAR }, { text: "42", style: { ...CHAR, sdtPath: ["field"] } }], sdtPath: ["sec"] },
      ],
      sdts: {
        sec: { type: "richText", tag: "Section", alias: "Sec" },
        field: { type: "plainText", tag: "Field" },
        agree: { type: "checkbox", tag: "Agree", checked: true },
      },
    };
    const rows = querySdts(doc);
    const byId = new Map(rows.map((r) => [r.id, r]));

    expect(byId.get("sec")).toMatchObject({
      sdtType: "richText", tag: "Section", alias: "Sec", checked: null, placeholder: false,
      parentId: null, childIds: ["field"], path: ["sec"], depth: 0, text: "Label: 42",
    });
    expect(byId.get("field")).toMatchObject({
      sdtType: "plainText", tag: "Field", parentId: "sec", path: ["sec", "field"], depth: 1, text: "42",
    });
    // A declared-but-unplaced checkbox: root, no members, checked reflected.
    expect(byId.get("agree")).toMatchObject({ sdtType: "checkbox", checked: true, parentId: null, childIds: [], text: "" });
  });

  it("is empty when the document has no controls", () => {
    expect(querySdts({ section: section(), blocks: [para("p1", "plain")] })).toEqual([]);
  });
});

describe("queryBridge — fields / bookmarks / notes / lists / styles / location / value", () => {
  const numberList = (): import("@cw/shared").ListDefinition => ({
    id: "L",
    levels: Array.from({ length: 9 }, (_, i) => ({ format: "decimal" as const, text: `%${i + 1}.`, indentLeftPx: 0, hangingPx: 0, start: 1 })),
  });

  it("queryFields flattens field defs", () => {
    const doc: Document = {
      section: section(),
      blocks: [para("p", "x")],
      fields: { f1: { id: "f1", instruction: " PAGE ", name: "PAGE", kind: "builtin" } },
    };
    expect(queryFields(doc)).toEqual([{ id: "f1", name: "PAGE", kind: "builtin", instruction: " PAGE " }]);
  });

  it("queryBookmarks flattens ranges", () => {
    const doc: Document = {
      section: section(),
      blocks: [para("p", "hello")],
      bookmarks: { bm: { start: { blockId: "p", offset: 0 }, end: { blockId: "p", offset: 5 } } },
    };
    expect(queryBookmarks(doc)).toEqual([{ name: "bm", startBlockId: "p", startOffset: 0, endBlockId: "p", endOffset: 5 }]);
  });

  it("queryFootnotes / queryEndnotes flatten note stories to id + text", () => {
    const doc: Document = {
      section: section(),
      blocks: [para("p", "body")],
      footnotes: { fn1: [para("fp", "foot")] },
      endnotes: { en1: [para("ep", "end")] },
    };
    expect(queryFootnotes(doc)).toEqual([{ id: "fn1", text: "foot" }]);
    expect(queryEndnotes(doc)).toEqual([{ id: "en1", text: "end" }]);
  });

  it("queryListItems resolves markers with container", () => {
    const li = (id: string, level: number): Paragraph => ({ ...para(id, id), style: { ...PARA, list: { listId: "L", level } } });
    const doc: Document = { section: section(), blocks: [li("a", 0), li("b", 0)], lists: { L: numberList() } };
    expect(queryListItems(doc, "L")).toEqual([
      { paragraphId: "a", level: 0, marker: "1.", text: "a", container: "body" },
      { paragraphId: "b", level: 0, marker: "2.", text: "b", container: "body" },
    ]);
  });

  it("queryStyles flattens the stylesheet", () => {
    const doc: Document = {
      section: section(),
      blocks: [para("p", "x")],
      stylesheet: { defaultStyleId: "Normal", styles: [{ id: "H1", name: "Heading 1", type: "paragraph", char: {}, para: {} }] },
    };
    expect(queryStyles(doc)).toEqual([{ id: "H1", name: "Heading 1", styleType: "paragraph", basedOn: null }]);
  });

  it("queryBlockPath reports location (null for unknown)", () => {
    const doc: Document = { section: section(), blocks: [para("p", "x"), table("t", para("cp", "in cell"))] };
    expect(queryBlockPath(doc, "cp")).toEqual({ container: "body", tableId: "t", row: 0, col: 0, noteKind: null, noteId: null });
    expect(queryBlockPath(doc, "p")).toEqual({ container: "body", tableId: null, row: -1, col: -1, noteKind: null, noteId: null });
    expect(queryBlockPath(doc, "nope")).toBeNull();
  });

  it("queryPositionOfText finds the first offset (null if absent)", () => {
    const doc: Document = { section: section(), blocks: [para("p", "hello world")] };
    expect(queryPositionOfText(doc, "world")).toEqual({ blockId: "p", offset: 6 });
    expect(queryPositionOfText(doc, "nope")).toBeNull();
  });

  it("queryRangeText slices across positions", () => {
    const doc: Document = { section: section(), blocks: [para("a", "hello world"), para("b", "second")] };
    expect(queryRangeText(doc, "a", 6, "b", 2)).toBe("world\nse");
  });

  it("querySdtValue returns the typed value (null if absent)", () => {
    const doc: Document = {
      section: section(),
      blocks: [{ ...para("p", "One"), runs: [{ text: "One", style: { ...CHAR, sdtPath: ["c"] } }] }],
      sdts: { c: { type: "dropDown", listItems: [{ display: "One", value: "1" }] } },
    };
    expect(querySdtValue(doc, "c")).toEqual({ sdtType: "dropDown", text: "One", checked: null, selected: "1" });
    expect(querySdtValue(doc, "nope")).toBeNull();
  });
});

describe("queryBridge.mapPages", () => {
  it("flattens layout page margins", () => {
    const pages: LayoutPageInfo[] = [
      { index: 0, number: 1, widthPx: 816, heightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 }, contentTopPx: 96, contentBottomPx: 960, blockIds: ["a", "b"] },
    ];
    expect(mapPages(pages)).toEqual([
      { index: 0, number: 1, widthPx: 816, heightPx: 1056, marginTop: 96, marginRight: 96, marginBottom: 96, marginLeft: 96, blockIds: ["a", "b"] },
    ]);
  });
});
