import { describe, expect, it } from "vitest";
import { isBuiltinField, isCustomFieldInstruction, parseFieldInstruction } from "./fields";

describe("parseFieldInstruction", () => {
  it("splits keyword + quoted args", () => {
    const p = parseFieldInstruction(' MYCHART "sales-2026" ');
    expect(p.name).toBe("MYCHART");
    expect(p.args).toEqual(["sales-2026"]);
    expect(p.raw).toBe(' MYCHART "sales-2026" ');
  });

  it("uppercases the keyword and keeps switches as args", () => {
    const p = parseFieldInstruction(' toc \\o "1-3" \\h ');
    expect(p.name).toBe("TOC");
    expect(p.args).toEqual(["\\o", "1-3", "\\h"]);
  });

  it("handles a bare keyword and a blank instruction", () => {
    expect(parseFieldInstruction(" PAGE ").name).toBe("PAGE");
    expect(parseFieldInstruction("   ").name).toBe("");
  });
});

describe("isBuiltinField", () => {
  it("accepts built-ins (case-insensitive) and rejects custom names", () => {
    for (const n of ["TOC", "page", "PageRef", "HYPERLINK", "ref"]) expect(isBuiltinField(n)).toBe(true);
    for (const n of ["MYCHART", "WIDGET", "salesTable"]) expect(isBuiltinField(n)).toBe(false);
  });
});

describe("isCustomFieldInstruction", () => {
  it("is true only for non-built-in keyworded instructions", () => {
    expect(isCustomFieldInstruction(' MYCHART "x" ')).toBe(true);
    expect(isCustomFieldInstruction(' TOC \\o "1-3" ')).toBe(false);
    expect(isCustomFieldInstruction(" PAGEREF _Toc1 \\h ")).toBe(false);
    expect(isCustomFieldInstruction("   ")).toBe(false);
  });
});
