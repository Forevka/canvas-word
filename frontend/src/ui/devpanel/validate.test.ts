import { describe, it, expect } from "vitest";
import type { CharStyle, Document, ParaStyle, Paragraph, Run } from "@cw/shared";
import { countBySeverity, validateDocument } from "./validate";

const CS: CharStyle = { fontFamily: "x", fontSizePx: 12, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PS: ParaStyle = { align: "left", lineHeight: 1, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const run = (text: string, style: Partial<CharStyle> = {}): Run => ({ text, style: { ...CS, ...style } });
const para = (id: string, runs: Run[] = [], style: Partial<ParaStyle> = {}, extra: Partial<Paragraph> = {}): Paragraph =>
  ({ kind: "paragraph", id, revision: 0, runs, style: { ...PS, ...style }, ...extra });
const doc = (partial: Partial<Document>): Document => ({ section: {} as Document["section"], blocks: [], ...partial });

const codes = (d: Document): string[] => validateDocument(d).map((x) => x.code);

describe("validateDocument", () => {
  it("passes a clean document", () => {
    expect(validateDocument(doc({ blocks: [para("a", [run("hi")])] }))).toEqual([]);
  });

  it("flags a dangling content-control reference (block + run)", () => {
    const d = doc({ blocks: [para("a", [run("x", { sdtPath: ["s9"] })], {}, { sdtPath: ["s8"] })] });
    expect(codes(d)).toContain("sdt-missing");
    expect(validateDocument(d).filter((x) => x.code === "sdt-missing")).toHaveLength(2);
  });

  it("flags an unused content control", () => {
    const d = doc({ blocks: [para("a", [run("x")])], sdts: { s1: { type: "richText" } } });
    expect(codes(d)).toContain("sdt-unused");
  });

  it("flags a dangling field and a duplicate block id", () => {
    const d = doc({ blocks: [para("a", [run("x", { fieldId: "f1" })]), para("a", [run("y")])] });
    expect(codes(d)).toContain("field-missing");
    expect(codes(d)).toContain("duplicate-id");
  });

  it("flags missing list / named style / table style references", () => {
    const d = doc({
      blocks: [
        para("a", [run("x")], { list: { listId: "L1", level: 0 }, namedStyle: "Heading1" }),
        { kind: "table", id: "t", revision: 0, rows: [], styleId: "TS1" },
      ],
    });
    const c = codes(d);
    expect(c).toContain("list-missing");
    expect(c).toContain("style-missing");
    expect(c).toContain("tablestyle-missing");
  });

  it("flags a bookmark anchored to a missing block", () => {
    const d = doc({ blocks: [para("a", [run("x")])], bookmarks: { bm: { start: { blockId: "ghost", offset: 0 }, end: { blockId: "ghost", offset: 1 } } } });
    const c = codes(d);
    expect(c).toContain("bookmark-start-missing");
    expect(c).toContain("bookmark-end-missing");
  });

  it("counts diagnostics by severity", () => {
    const d = doc({ blocks: [para("a", [run("x", { fieldId: "f1" })])], sdts: { s1: { type: "richText" } } });
    const c = countBySeverity(validateDocument(d));
    expect(c.error).toBeGreaterThanOrEqual(1); // field-missing
    expect(c.info).toBeGreaterThanOrEqual(1); // sdt-unused
  });
});
