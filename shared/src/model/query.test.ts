import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, SectionProps, TableBlock } from "./document";
import {
  findParagraphs,
  getBlockById,
  getImageById,
  getParagraphById,
  getParagraphs,
  getSdt,
  getSdtAncestors,
  getSdtBlocks,
  getSdtChildren,
  getSdtDescendants,
  getSdtNodes,
  getSdtRoots,
  getSdts,
  getSdtsByAlias,
  getSdtsByTag,
  getSdtValue,
  getSections,
  getTableById,
  getTables,
  sdtText,
  textOf,
  walk,
  getField,
  getFields,
  getFieldsByName,
  getFieldBlocks,
  getBookmark,
  getBookmarks,
  getFootnotes,
  getEndnotes,
  getListItems,
  getStyles,
  getStyleById,
  blockPath,
} from "./query";
import type { Run } from "./document";
import { numberListDefinition } from "./lists";

const CHAR: CharStyle = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA = { align: "left" as const, lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };

const para = (id: string, text: string, extra: Partial<Paragraph["style"]> = {}): Paragraph => ({
  kind: "paragraph",
  id,
  revision: 0,
  runs: [{ text, style: CHAR }],
  style: { ...PARA, ...extra },
});

const cell = (id: string, ...blocks: Paragraph[]) => ({ id, blocks });

const table = (id: string, ...paras: Paragraph[]): TableBlock => ({
  kind: "table",
  id,
  revision: 0,
  rows: [{ cells: [cell(`${id}-c0`, ...paras)] }],
});

const section = (): SectionProps => ({ pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } });

/** body: p1, table(t1 → cellPara), p2; header band: h1; one footnote body. */
const richDoc = (): Document => ({
  section: { ...section(), header: [para("h1", "header text")] },
  blocks: [para("p1", "hello world"), table("t1", para("cellPara", "inside the cell")), para("p2", "goodbye world")],
  footnotes: { fn1: [para("fnp", "footnote body")] },
});

describe("query.walk", () => {
  it("descends into body, cells, bands, and notes by default", () => {
    const ids: string[] = [];
    walk(richDoc(), (b) => ids.push(b.id));
    expect(ids).toEqual(["p1", "t1", "cellPara", "p2", "h1", "fnp"]);
  });

  it("respects opts to skip cells, bands, and notes", () => {
    const ids: string[] = [];
    walk(richDoc(), (b) => ids.push(b.id), { cells: false, bands: false, notes: false });
    expect(ids).toEqual(["p1", "t1", "p2"]);
  });

  it("reports cell context for nested blocks", () => {
    let ctx: unknown;
    walk(richDoc(), (b, c) => {
      if (b.id === "cellPara") ctx = c;
    });
    expect(ctx).toEqual({ container: "body", cell: { tableId: "t1", row: 0, col: 0 } });
  });

  it("reports band container and note membership", () => {
    const seen: Record<string, unknown> = {};
    walk(richDoc(), (b, c) => {
      seen[b.id] = c;
    });
    expect(seen["h1"]).toEqual({ container: "header" });
    expect(seen["fnp"]).toEqual({ container: "body", note: { kind: "footnote", id: "fn1" } });
  });
});

describe("query.getParagraphs / getTables", () => {
  it("collects every paragraph across stories", () => {
    expect(getParagraphs(richDoc()).map((p) => p.id)).toEqual(["p1", "cellPara", "p2", "h1", "fnp"]);
  });

  it("body-only paragraphs when bands/notes/cells excluded", () => {
    expect(getParagraphs(richDoc(), { bands: false, notes: false, cells: false }).map((p) => p.id)).toEqual(["p1", "p2"]);
  });

  it("collects tables", () => {
    expect(getTables(richDoc()).map((t) => t.id)).toEqual(["t1"]);
  });
});

describe("query.textOf", () => {
  it("returns paragraph run text", () => {
    expect(textOf(para("x", "abc"))).toBe("abc");
  });

  it("joins table cells with tabs and rows with newlines", () => {
    const t: TableBlock = {
      kind: "table",
      id: "t",
      revision: 0,
      rows: [
        { cells: [cell("a", para("a1", "A")), cell("b", para("b1", "B"))] },
        { cells: [cell("c", para("c1", "C"))] },
      ],
    };
    expect(textOf(t)).toBe("A\tB\nC");
  });
});

describe("query.findParagraphs", () => {
  it("matches by substring across cells, bands, and notes", () => {
    const hits = findParagraphs(richDoc(), "world");
    expect(hits.map((h) => h.paragraph.id)).toEqual(["p1", "p2"]);
  });

  it("finds text inside a table cell", () => {
    const hits = findParagraphs(richDoc(), "inside");
    expect(hits).toHaveLength(1);
    expect(hits[0]!.context.cell).toEqual({ tableId: "t1", row: 0, col: 0 });
  });

  it("matches by RegExp and is not corrupted by a global flag", () => {
    const hits = findParagraphs(richDoc(), /World/gi);
    expect(hits.map((h) => h.paragraph.id)).toEqual(["p1", "p2"]);
  });
});

describe("query.getBlockById + typed getters", () => {
  it("finds a nested block by id", () => {
    expect(getBlockById(richDoc(), "cellPara")?.id).toBe("cellPara");
  });

  it("typed getters narrow by kind", () => {
    const doc = richDoc();
    expect(getParagraphById(doc, "p1")?.kind).toBe("paragraph");
    expect(getParagraphById(doc, "t1")).toBeUndefined(); // t1 is a table
    expect(getTableById(doc, "t1")?.kind).toBe("table");
    expect(getImageById(doc, "p1")).toBeUndefined();
  });
});

describe("query.walk — section-break bands & nested notes", () => {
  it("visits a header stored on an earlier section's section-break props", () => {
    const doc: Document = {
      section: section(),
      blocks: [
        para("a", "sec one", { sectionBreak: { type: "nextPage", props: { header: [para("sbHeader", "break header")] } } }),
        para("b", "sec two"),
      ],
    };
    const seen: Record<string, unknown> = {};
    walk(doc, (block, ctx) => {
      seen[block.id] = ctx;
    });
    expect(seen["sbHeader"]).toEqual({ container: "header" });
    expect(getParagraphs(doc).map((p) => p.id)).toContain("sbHeader");
  });

  it("keeps the band container for a table nested inside a header", () => {
    // The context spread also carries `note` into cells; note bodies are
    // paragraph-only in the model, so `container` is the reachable field to pin.
    const doc: Document = {
      section: { ...section(), header: [table("htbl", para("hcell", "head cell"))] },
      blocks: [para("p1", "body")],
    };
    let ctx: unknown;
    walk(doc, (block, c) => {
      if (block.id === "hcell") ctx = c;
    });
    expect(ctx).toEqual({ container: "header", cell: { tableId: "htbl", row: 0, col: 0 } });
  });
});

describe("query.textOf — multi-paragraph cell", () => {
  it("joins paragraphs within a cell with newlines", () => {
    const t: TableBlock = {
      kind: "table",
      id: "t",
      revision: 0,
      rows: [{ cells: [cell("c0", para("a", "A"), para("b", "B"))] }],
    };
    expect(textOf(t)).toBe("A\nB");
  });
});

describe("query.getSections", () => {
  it("returns one section for a plain document with the full block range", () => {
    const doc = richDoc();
    const secs = getSections(doc);
    expect(secs).toHaveLength(1);
    expect(secs[0]!.startBlock).toBe(0);
    expect(secs[0]!.endBlock).toBe(doc.blocks.length - 1);
  });

  it("splits at a section-break paragraph with per-section block ranges", () => {
    const doc: Document = {
      section: section(),
      blocks: [
        para("a", "sec one", { sectionBreak: { type: "nextPage", props: {} } }),
        para("b", "sec two body"),
      ],
    };
    const secs = getSections(doc);
    expect(secs).toHaveLength(2);
    expect(secs[0]).toMatchObject({ index: 0, startBlock: 0, endBlock: 0, breakType: "nextPage" });
    expect(secs[1]).toMatchObject({ index: 1, startBlock: 1, endBlock: 1 });
  });
});

// --- SDTs (content controls) — the primary templating surface -----------------

const sdtRun = (text: string, sdtPath?: string[]): Run => ({ text, style: sdtPath ? { ...CHAR, sdtPath } : CHAR });

const mkPara = (id: string, runs: Run[], extra: Partial<Paragraph> = {}): Paragraph => ({
  kind: "paragraph",
  id,
  revision: 0,
  runs,
  style: { ...PARA },
  ...extra,
});

describe("query SDT: flat lookups", () => {
  const doc = (): Document => ({
    section: section(),
    blocks: [mkPara("para", [sdtRun("before "), sdtRun("Ada", ["name"]), sdtRun(" after")])],
    sdts: {
      name: { type: "plainText", tag: "Name", alias: "Full Name" },
      note: { type: "richText", tag: "Name" }, // shares a tag with `name`
    },
  });

  it("getSdt returns props by id, undefined otherwise", () => {
    expect(getSdt(doc(), "name")).toMatchObject({ type: "plainText", tag: "Name" });
    expect(getSdt(doc(), "missing")).toBeUndefined();
    expect(getSdt({ section: section(), blocks: [] }, "x")).toBeUndefined();
  });

  it("getSdts lists every control", () => {
    expect(getSdts(doc()).map((s) => s.id).sort()).toEqual(["name", "note"]);
  });

  it("getSdtsByTag returns every control sharing the tag", () => {
    expect(getSdtsByTag(doc(), "Name").map((s) => s.id).sort()).toEqual(["name", "note"]);
    expect(getSdtsByTag(doc(), "nope")).toEqual([]);
  });

  it("getSdtsByAlias filters by title/alias", () => {
    expect(getSdtsByAlias(doc(), "Full Name").map((s) => s.id)).toEqual(["name"]);
  });

  it("sdtText reads an inline control's text", () => {
    expect(sdtText(doc(), "name")).toBe("Ada");
  });
});

describe("query SDT: nested (block-level control wrapping an inline control)", () => {
  // `sec` is a block-level control over two paragraphs; `field` is an inline
  // control on a run WITHIN the first paragraph → nested under `sec`.
  const doc = (): Document => ({
    section: section(),
    blocks: [
      mkPara("b0", [sdtRun("Label: "), sdtRun("42", ["field"])], { sdtPath: ["sec"] }),
      mkPara("b1", [sdtRun("more")], { sdtPath: ["sec"] }),
    ],
    sdts: {
      sec: { type: "richText", tag: "Section" },
      field: { type: "plainText", tag: "Field" },
    },
  });

  it("derives roots / children / ancestors / descendants", () => {
    expect(getSdtRoots(doc()).map((n) => n.id)).toEqual(["sec"]);
    expect(getSdtChildren(doc(), "sec").map((n) => n.id)).toEqual(["field"]);
    expect(getSdtChildren(doc(), "field")).toEqual([]);
    expect(getSdtAncestors(doc(), "field").map((n) => n.id)).toEqual(["sec"]);
    expect(getSdtAncestors(doc(), "sec")).toEqual([]);
    expect(getSdtDescendants(doc(), "sec").map((n) => n.id)).toEqual(["field"]);
  });

  it("nodes carry parentId / path / depth / childIds", () => {
    const byId = new Map(getSdtNodes(doc()).map((n) => [n.id, n]));
    expect(byId.get("sec")).toMatchObject({ path: ["sec"], depth: 0, childIds: ["field"] });
    expect(byId.get("sec")!.parentId).toBeUndefined();
    expect(byId.get("field")).toMatchObject({ parentId: "sec", path: ["sec", "field"], depth: 1, childIds: [] });
  });

  it("getSdtBlocks returns block-level members only", () => {
    expect(getSdtBlocks(doc(), "sec").map((b) => b.id)).toEqual(["b0", "b1"]);
    expect(getSdtBlocks(doc(), "field")).toEqual([]); // inline control has no block members
  });

  it("sdtText: block-level joins its blocks; inline reads its runs", () => {
    expect(sdtText(doc(), "sec")).toBe("Label: 42\nmore");
    expect(sdtText(doc(), "field")).toBe("42");
  });
});

describe("query SDT: deeply nested inline controls", () => {
  // One run belongs to control `b` nested inside `a` (inline path a→b).
  const doc = (): Document => ({
    section: section(),
    blocks: [mkPara("p", [sdtRun("x"), sdtRun("y", ["a", "b"]), sdtRun("z")])],
    sdts: { a: { type: "richText" }, b: { type: "plainText" } },
  });

  it("reads nesting from a run's full inline path", () => {
    expect(getSdtRoots(doc()).map((n) => n.id)).toEqual(["a"]);
    expect(getSdtChildren(doc(), "a").map((n) => n.id)).toEqual(["b"]);
    expect(getSdtAncestors(doc(), "b").map((n) => n.id)).toEqual(["a"]);
    expect(getSdtNodes(doc()).find((n) => n.id === "b")).toMatchObject({ parentId: "a", path: ["a", "b"], depth: 1 });
  });

  it("sdtText includes descendants (membership is by path containment)", () => {
    expect(sdtText(doc(), "a")).toBe("y");
    expect(sdtText(doc(), "b")).toBe("y");
  });
});

describe("query SDT: controls in cells and header bands are traversed", () => {
  const doc = (): Document => ({
    section: { ...section(), header: [mkPara("h", [sdtRun("H "), sdtRun("hv", ["hdr"])])] },
    blocks: [table("t", mkPara("cp", [sdtRun("cell "), sdtRun("cv", ["cell"])]))],
    sdts: { cell: { type: "plainText", tag: "Cell" }, hdr: { type: "plainText", tag: "Hdr" } },
  });

  it("finds inline controls inside table cells and header bands", () => {
    expect(sdtText(doc(), "cell")).toBe("cv");
    expect(sdtText(doc(), "hdr")).toBe("hv");
    expect(getSdtRoots(doc()).map((n) => n.id).sort()).toEqual(["cell", "hdr"]);
  });
});

describe("query SDT: edge cases", () => {
  it("a declared control with no membership path is a childless root with empty text", () => {
    const doc: Document = {
      section: section(),
      blocks: [mkPara("p", [sdtRun("plain")])],
      sdts: { orphan: { type: "checkbox", tag: "Agree", checked: true } },
    };
    expect(getSdtNodes(doc)[0]).toMatchObject({ id: "orphan", childIds: [], path: ["orphan"], depth: 0 });
    expect(getSdtNodes(doc)[0]!.parentId).toBeUndefined();
    expect(sdtText(doc, "orphan")).toBe("");
    expect(getSdt(doc, "orphan")?.checked).toBe(true); // checkbox value lives in props
  });

  it("no sdts → empty everywhere; unknown ids → empty", () => {
    const doc: Document = { section: section(), blocks: [mkPara("p", [sdtRun("x")])] };
    expect(getSdts(doc)).toEqual([]);
    expect(getSdtNodes(doc)).toEqual([]);
    expect(getSdtRoots(doc)).toEqual([]);
    expect(getSdtChildren(doc, "x")).toEqual([]);
    expect(getSdtAncestors(doc, "x")).toEqual([]);
    expect(getSdtDescendants(doc, "x")).toEqual([]);
    expect(getSdtBlocks(doc, "x")).toEqual([]);
    expect(sdtText(doc, "x")).toBe("");
    expect(getSdtValue(doc, "x")).toBeUndefined();
  });
});

describe("query SDT: getSdtValue (typed value read)", () => {
  it("reads text controls as plain text", () => {
    const doc: Document = {
      section: section(),
      blocks: [mkPara("p", [sdtRun("Ada", ["name"])])],
      sdts: { name: { type: "plainText", tag: "Name" } },
    };
    expect(getSdtValue(doc, "name")).toEqual({ type: "plainText", text: "Ada" });
  });

  it("reads a checkbox's state", () => {
    const doc: Document = {
      section: section(),
      blocks: [mkPara("p", [sdtRun("x")])],
      sdts: { agree: { type: "checkbox", checked: true } },
    };
    expect(getSdtValue(doc, "agree")).toEqual({ type: "checkbox", text: "", checked: true });
  });

  it("resolves a dropdown's selected VALUE from its display text", () => {
    const doc: Document = {
      section: section(),
      blocks: [mkPara("p", [sdtRun("Two", ["choice"])])],
      sdts: { choice: { type: "dropDown", listItems: [{ display: "One", value: "1" }, { display: "Two", value: "2" }] } },
    };
    expect(getSdtValue(doc, "choice")).toEqual({ type: "dropDown", text: "Two", selected: "2" });
  });

  it("falls back to the raw text for a combo box's free entry", () => {
    const doc: Document = {
      section: section(),
      blocks: [mkPara("p", [sdtRun("Custom", ["combo"])])],
      sdts: { combo: { type: "comboBox", listItems: [{ display: "Alpha", value: "a" }] } },
    };
    expect(getSdtValue(doc, "combo")).toEqual({ type: "comboBox", text: "Custom", selected: "Custom" });
  });

  it("leaves selected undefined for a dropdown whose text matches no option", () => {
    // A malformed/imported dropdown whose current text is not a listItem display.
    const doc: Document = {
      section: section(),
      blocks: [mkPara("p", [sdtRun("Stale", ["choice"])])],
      sdts: { choice: { type: "dropDown", listItems: [{ display: "One", value: "1" }] } },
    };
    const v = getSdtValue(doc, "choice")!;
    expect(v).toEqual({ type: "dropDown", text: "Stale" }); // no `selected` key
    expect(v.selected).toBeUndefined();
  });
});

describe("query: fields", () => {
  const doc = (): Document => ({
    section: section(),
    blocks: [{ ...para("b", "chart"), fieldId: "f2" }],
    fields: {
      f1: { id: "f1", instruction: " PAGE ", name: "PAGE", kind: "builtin" },
      f2: { id: "f2", instruction: ' MYCHART "x" ', name: "MYCHART", kind: "custom" },
    },
  });

  it("looks up by id / all / by name (case-insensitive)", () => {
    expect(getField(doc(), "f1")?.name).toBe("PAGE");
    expect(getField(doc(), "nope")).toBeUndefined();
    expect(getFields(doc()).map((f) => f.id).sort()).toEqual(["f1", "f2"]);
    expect(getFieldsByName(doc(), "page").map((f) => f.id)).toEqual(["f1"]);
  });

  it("getFieldBlocks returns the region blocks", () => {
    expect(getFieldBlocks(doc(), "f2").map((b) => b.id)).toEqual(["b"]);
    expect(getFieldBlocks(doc(), "f1")).toEqual([]);
  });
});

describe("query: bookmarks", () => {
  const doc: Document = {
    section: section(),
    blocks: [para("p", "hello")],
    bookmarks: { bm: { start: { blockId: "p", offset: 0 }, end: { blockId: "p", offset: 3 } } },
  };

  it("returns a range by name and enumerates entries", () => {
    expect(getBookmark(doc, "bm")).toEqual({ start: { blockId: "p", offset: 0 }, end: { blockId: "p", offset: 3 } });
    expect(getBookmark(doc, "x")).toBeUndefined();
    expect(getBookmarks(doc).map((b) => b.name)).toEqual(["bm"]);
  });
});

describe("query: notes", () => {
  const doc: Document = {
    section: section(),
    blocks: [para("p", "body")],
    footnotes: { fn1: [para("fp", "foot")] },
    endnotes: { en1: [para("ep", "end")] },
  };

  it("enumerates footnote and endnote stories", () => {
    expect(getFootnotes(doc)).toEqual([{ id: "fn1", paragraphs: [para("fp", "foot")] }]);
    expect(getEndnotes(doc).map((n) => n.id)).toEqual(["en1"]);
    expect(getFootnotes({ section: section(), blocks: [] })).toEqual([]);
  });
});

describe("query: getListItems", () => {
  it("resolves markers in body reading order, resetting deeper levels", () => {
    const li = (id: string, level: number): Paragraph => ({ ...para(id, id), style: { ...PARA, list: { listId: "L", level } } });
    const doc: Document = {
      section: section(),
      blocks: [li("a", 0), li("b", 1), li("c", 0), li("d", 1)],
      lists: { L: numberListDefinition("L", "decimal", ".") },
    };
    const items = getListItems(doc, "L");
    expect(items.map((i) => [i.paragraph.id, i.level, i.marker])).toEqual([
      ["a", 0, "1."],
      ["b", 1, "1."],
      ["c", 0, "2."], // back to level 0 → level-1 counter resets
      ["d", 1, "1."], // level-1 restarts at 1, proving the reset
    ]);
    expect(getListItems(doc, "other")).toEqual([]);
  });
});

describe("query: styles", () => {
  const doc: Document = {
    section: section(),
    blocks: [para("p", "x")],
    stylesheet: { defaultStyleId: "Normal", styles: [{ id: "H1", name: "Heading 1", char: {}, para: {} }] },
  };

  it("enumerates and looks up named styles", () => {
    expect(getStyles(doc).map((s) => s.id)).toEqual(["H1"]);
    expect(getStyleById(doc, "H1")?.name).toBe("Heading 1");
    expect(getStyleById(doc, "nope")).toBeUndefined();
    expect(getStyles({ section: section(), blocks: [] })).toEqual([]);
  });
});

describe("query: blockPath", () => {
  it("reports the container/cell/note context of a block", () => {
    const doc = richDoc();
    expect(blockPath(doc, "p1")).toEqual({ container: "body" });
    expect(blockPath(doc, "cellPara")).toEqual({ container: "body", cell: { tableId: "t1", row: 0, col: 0 } });
    expect(blockPath(doc, "h1")).toEqual({ container: "header" });
    expect(blockPath(doc, "fnp")).toEqual({ container: "body", note: { kind: "footnote", id: "fn1" } });
    expect(blockPath(doc, "nope")).toBeUndefined();
  });
});
