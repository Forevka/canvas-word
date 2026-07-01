import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, SectionProps, TableBlock } from "./document";
import {
  findParagraphs,
  getBlockById,
  getImageById,
  getParagraphById,
  getParagraphs,
  getSections,
  getTableById,
  getTables,
  textOf,
  walk,
} from "./query";

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
