// Custom-field export tested by ROUND-TRIP: a custom field imported from a docx
// must re-export as a real complex field (begin/instrText/separate … end) and
// re-import to the same field — instruction preserved, result blocks preserved,
// and stable across a second round-trip.

import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import type { Document, Paragraph } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import { simpleDocx } from "../../import/docx/fixture";
import { writeDocx } from "./writeDocx";

const fld = (t: string): string => `<w:r><w:fldChar w:fldCharType="${t}"/></w:r>`;
const instr = (s: string): string => `<w:r><w:instrText xml:space="preserve">${s}</w:instrText></w:r>`;
const txt = (s: string): string => `<w:r><w:t>${s}</w:t></w:r>`;
const roundTrip = (doc: Document): Document => runImport(writeDocx(doc).bytes).doc;
const docXml = (doc: Document): string => strFromU8(unzipSync(writeDocx(doc).bytes)["word/document.xml"]!);
const fieldDefs = (d: Document) => Object.values(d.fields ?? {});

describe("custom field export — round trip", () => {
  const SINGLE = `<w:p>${fld("begin")}${instr(' MYCHART "sales" ')}${fld("separate")}${txt("[chart]")}${fld("end")}</w:p>`;

  it("re-emits a single-paragraph field as a live complex field", () => {
    const a = runImport(simpleDocx(SINGLE)).doc;
    const xml = docXml(a);
    expect(xml).toContain('<w:instrText xml:space="preserve"> MYCHART "sales" </w:instrText>');
    expect(xml).toContain('w:fldCharType="begin"');
    expect(xml).toContain('w:fldCharType="separate"');
    expect(xml).toContain('w:fldCharType="end"');

    const b = roundTrip(a);
    expect(fieldDefs(b).length).toBe(1);
    expect(fieldDefs(b)[0]!.instruction).toBe(' MYCHART "sales" ');
    const p = b.blocks.find((x): x is Paragraph => x.kind === "paragraph")!;
    expect(p.fieldId).toBe(fieldDefs(b)[0]!.id);
    expect(p.runs.map((r) => r.text).join("")).toBe("[chart]");
  });

  it("is idempotent across a second round trip (no drift)", () => {
    const a = runImport(simpleDocx(SINGLE)).doc;
    const b = roundTrip(a);
    const c = roundTrip(b);
    expect(fieldDefs(c).length).toBe(1);
    expect(fieldDefs(c)[0]!.instruction).toBe(' MYCHART "sales" ');
    // same number of body paragraphs both times — no synthetic blocks accreting
    const bp = b.blocks.filter((x) => x.kind === "paragraph").length;
    const cp = c.blocks.filter((x) => x.kind === "paragraph").length;
    expect(cp).toBe(bp);
  });

  it("wraps a multi-block field result (paragraphs) in one field region", () => {
    const body =
      `<w:p>${fld("begin")}${instr(" WIDGET ")}${fld("separate")}${txt("L1")}</w:p>` +
      `<w:p>${txt("L2")}</w:p>` +
      `<w:p>${txt("L3")}${fld("end")}</w:p>`;
    const a = runImport(simpleDocx(body)).doc;
    const xml = docXml(a);
    // Exactly one begin/separate/end pair around the three result paragraphs.
    expect((xml.match(/w:fldCharType="begin"/g) ?? []).length).toBe(1);
    expect((xml.match(/w:fldCharType="end"/g) ?? []).length).toBe(1);

    const b = roundTrip(a);
    const id = fieldDefs(b)[0]!.id;
    const ps = b.blocks.filter((x): x is Paragraph => x.kind === "paragraph");
    expect(ps.filter((p) => p.fieldId === id).length).toBe(3);
    expect(ps.map((p) => p.runs.map((r) => r.text).join(""))).toEqual(["L1", "L2", "L3"]);
  });
});
