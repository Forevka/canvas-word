// Custom (non-built-in) field preservation on import: the result blocks are
// tagged with a fieldId and the verbatim instruction is registered in
// Document.fields, while the cached result content is kept (renders until the
// host refreshes it). Built-in fields (DATE, PAGEREF, …) are unaffected.

import { describe, expect, it } from "vitest";
import type { Paragraph } from "@cw/shared";
import { runImport } from "./pipeline";
import { simpleDocx } from "./fixture";

const fld = (type: "begin" | "separate" | "end"): string =>
  `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
const instr = (s: string): string => `<w:r><w:instrText xml:space="preserve">${s}</w:instrText></w:r>`;
const txt = (s: string): string => `<w:r><w:t>${s}</w:t></w:r>`;
const paras = (blocks: { kind: string }[]): Paragraph[] =>
  blocks.filter((b): b is Paragraph => b.kind === "paragraph");
const text = (p: Paragraph): string => p.runs.map((r) => r.text).join("");

describe("custom field import", () => {
  it("captures a single-paragraph custom field + keeps its cached result", () => {
    const body = `<w:p>${fld("begin")}${instr(' MYCHART "sales-2026" ')}${fld("separate")}${txt("[chart]")}${fld("end")}</w:p>`;
    const doc = runImport(simpleDocx(body)).doc;

    expect(doc.fields).toBeDefined();
    const defs = Object.values(doc.fields!);
    expect(defs.length).toBe(1);
    expect(defs[0]!.name).toBe("MYCHART");
    expect(defs[0]!.instruction).toBe(' MYCHART "sales-2026" ');
    expect(defs[0]!.kind).toBe("custom");

    const p = paras(doc.blocks)[0]!;
    expect(p.fieldId).toBe(defs[0]!.id);
    expect(text(p)).toBe("[chart]"); // cached result preserved
  });

  it("tags every block of a field whose result spans multiple paragraphs", () => {
    const body =
      `<w:p>${fld("begin")}${instr(" WIDGET ")}${fld("separate")}${txt("Line1")}</w:p>` +
      `<w:p>${txt("Line2")}</w:p>` +
      `<w:p>${txt("Line3")}${fld("end")}</w:p>` +
      `<w:p>${txt("After")}</w:p>`;
    const doc = runImport(simpleDocx(body)).doc;

    const id = Object.values(doc.fields!)[0]!.id;
    const ps = paras(doc.blocks);
    expect(ps.slice(0, 3).every((p) => p.fieldId === id)).toBe(true);
    expect(text(ps[3]!)).toBe("After");
    expect(ps[3]!.fieldId).toBeUndefined(); // outside the field
  });

  it("does NOT capture built-in fields (DATE) as custom fields", () => {
    const body = `<w:p>${fld("begin")}${instr(" DATE \\@ \"M/d/yyyy\" ")}${fld("separate")}${txt("1/1/2026")}${fld("end")}</w:p>`;
    const doc = runImport(simpleDocx(body)).doc;

    expect(doc.fields).toBeUndefined();
    const p = paras(doc.blocks)[0]!;
    expect(p.fieldId).toBeUndefined();
    expect(text(p)).toBe("1/1/2026");
  });
});
