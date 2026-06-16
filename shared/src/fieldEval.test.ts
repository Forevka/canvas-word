import { describe, expect, it } from "vitest";
import type { CharStyle, FieldSpec } from "./model/document";
import { buildInstruction, evaluateField, evaluateIf, formatFieldDate, parseFieldSpec } from "./fieldEval";
import { parseFieldInstruction } from "./fields";

const CHAR: CharStyle = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const ctx = { now: new Date(2026, 5, 16, 14, 9, 7) }; // 2026-06-16 14:09:07 (Tue)

describe("formatFieldDate", () => {
  it("formats a Word \\@ picture", () => {
    expect(formatFieldDate(ctx.now, "M/d/yyyy")).toBe("6/16/2026");
    expect(formatFieldDate(ctx.now, "MM/dd/yy")).toBe("06/16/26");
    expect(formatFieldDate(ctx.now, "dddd, MMMM d, yyyy")).toBe("Tuesday, June 16, 2026");
    expect(formatFieldDate(ctx.now, "h:mm AM/PM")).toBe("2:09 PM");
    expect(formatFieldDate(ctx.now, "HH:mm:ss")).toBe("14:09:07");
    expect(formatFieldDate(ctx.now, "'Year' yyyy")).toBe("Year 2026");
  });
});

describe("evaluateField", () => {
  it("PAGE/NUMPAGES → tokens (with format suffix)", () => {
    expect(evaluateField({ type: "PAGE" }, CHAR, ctx)).toEqual({ kind: "token", token: "{page}" });
    expect(evaluateField({ type: "PAGE", numFmt: "roman" }, CHAR, ctx)).toEqual({ kind: "token", token: "{page:roman}" });
    expect(evaluateField({ type: "NUMPAGES" }, CHAR, ctx)).toEqual({ kind: "token", token: "{pages}" });
  });
  it("DATE → materialized run", () => {
    const r = evaluateField({ type: "DATE", format: "M/d/yyyy" }, CHAR, ctx);
    expect(r).toEqual({ kind: "runs", runs: [{ text: "6/16/2026", style: CHAR }] });
  });
});

describe("evaluateIf", () => {
  const ifSpec = (operandA: string, op: "=" | "<>" | "<" | ">" | "<=" | ">=", operandB: string): Extract<FieldSpec, { type: "IF" }> => ({
    type: "IF", operandA, op, operandB,
    trueRuns: [{ text: "YES", style: CHAR }], falseRuns: [{ text: "no", style: CHAR }],
  });
  it("numeric comparison", () => {
    expect(evaluateIf(ifSpec("2", ">", "1"))[0]!.text).toBe("YES");
    expect(evaluateIf(ifSpec("2", "<", "1"))[0]!.text).toBe("no");
    expect(evaluateIf(ifSpec("3", "=", "3"))[0]!.text).toBe("YES");
    expect(evaluateIf(ifSpec("3", "<>", "3"))[0]!.text).toBe("no");
  });
  it("string comparison when not both numeric", () => {
    expect(evaluateIf(ifSpec("apple", "<", "banana"))[0]!.text).toBe("YES");
    expect(evaluateIf(ifSpec("x", "=", "x"))[0]!.text).toBe("YES");
  });
  it("returns cloned runs (no spec mutation)", () => {
    const s = ifSpec("1", "=", "1");
    const out = evaluateIf(s);
    out[0]!.text = "MUT";
    expect(s.trueRuns[0]!.text).toBe("YES");
  });
});

describe("spec ⇄ instruction round-trip", () => {
  const round = (spec: FieldSpec): FieldSpec | undefined =>
    parseFieldSpec(parseFieldInstruction(buildInstruction(spec)), CHAR);

  it("PAGE with roman format", () => {
    expect(round({ type: "PAGE", numFmt: "roman" })).toEqual({ type: "PAGE", numFmt: "roman" });
  });
  it("DATE format", () => {
    expect(round({ type: "DATE", format: "MMMM d, yyyy" })).toEqual({ type: "DATE", format: "MMMM d, yyyy" });
  });
  it("IF operands + op + text branches", () => {
    const spec: FieldSpec = { type: "IF", operandA: "5", op: ">=", operandB: "3", trueRuns: [{ text: "big", style: CHAR }], falseRuns: [{ text: "small", style: CHAR }] };
    const r = round(spec);
    expect(r?.type).toBe("IF");
    if (r?.type === "IF") {
      expect([r.operandA, r.op, r.operandB]).toEqual(["5", ">=", "3"]);
      expect(r.trueRuns[0]!.text).toBe("big");
      expect(r.falseRuns[0]!.text).toBe("small");
    }
  });
  it("custom field → no spec", () => {
    expect(parseFieldSpec(parseFieldInstruction(' MYCHART "x" '), CHAR)).toBeUndefined();
  });
});
