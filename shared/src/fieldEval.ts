// Field evaluation — pure, layout-free, so import, the editor, and the headless
// render all agree. Two result shapes:
//   • "token"  — layout-dependent fields (PAGE/NUMPAGES). The result run holds a
//     `{page}`/`{pages}` token the layout engine re-substitutes per page (see
//     engine.ts substituteTokens). The run's CharStyle.fieldId is what marks it a
//     real field vs literal text.
//   • "runs"   — static fields (DATE/TIME/IF), materialized once on insert/update.
//
// `spec` ⇄ `instruction` round-trips via parseFieldSpec / buildInstruction; the
// verbatim instruction stays the export source of truth (unknown switches survive).

import type { CharStyle, FieldSpec, IfOp, PageNumFmt, Run } from "./model/document";
import { parseFieldInstruction, type ParsedFieldInstruction } from "./fields";
import { textOfRuns } from "./model/text";

export interface FieldEvalCtx {
  /** "Now" for DATE/TIME — passed in so evaluation stays pure/deterministic. */
  now: Date;
}

export type FieldEvalResult =
  | { kind: "runs"; runs: Run[] }
  | { kind: "token"; token: string };

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

const cloneRuns = (runs: Run[]): Run[] => runs.map((r) => ({ text: r.text, style: { ...r.style } }));

/** Token suffix for a page-number format (matches engine.ts `{page:fmt}` grammar). */
const numFmtToken = (f: PageNumFmt | undefined): string =>
  f && f !== "arabic" ? `:${f}` : "";

/** OOXML `\* <fmt>` switch for a page-number format (matches export fieldInstr). */
const numFmtSwitch = (f: PageNumFmt | undefined): string =>
  f === "roman" ? " \\* roman"
  : f === "Roman" ? " \\* ROMAN"
  : f === "alpha" ? " \\* alphabetic"
  : f === "Alpha" ? " \\* ALPHABETIC"
  : "";

/** Format a Date with a Word `\@` picture (subset): yyyy/yy, MMMM/MMM/MM/M,
 *  dddd/ddd/dd/d, HH/H, hh/h, mm/m, ss/s, AM/PM, am/pm. `'literal'` passes through. */
export function formatFieldDate(d: Date, fmt: string): string {
  const pad = (n: number, w = 2): string => String(n).padStart(w, "0");
  const h12 = ((d.getHours() + 11) % 12) + 1;
  let out = "";
  let i = 0;
  while (i < fmt.length) {
    if (fmt[i] === "'") {
      const end = fmt.indexOf("'", i + 1);
      if (end < 0) { out += fmt.slice(i + 1); break; }
      out += fmt.slice(i + 1, end);
      i = end + 1;
      continue;
    }
    const rest = fmt.slice(i);
    const tok = (t: string): boolean => rest.startsWith(t);
    if (tok("yyyy")) { out += pad(d.getFullYear(), 4); i += 4; }
    else if (tok("yy")) { out += pad(d.getFullYear() % 100); i += 2; }
    else if (tok("MMMM")) { out += MONTHS[d.getMonth()]; i += 4; }
    else if (tok("MMM")) { out += MONTHS[d.getMonth()]!.slice(0, 3); i += 3; }
    else if (tok("MM")) { out += pad(d.getMonth() + 1); i += 2; }
    else if (tok("M")) { out += String(d.getMonth() + 1); i += 1; }
    else if (tok("dddd")) { out += DAYS[d.getDay()]; i += 4; }
    else if (tok("ddd")) { out += DAYS[d.getDay()]!.slice(0, 3); i += 3; }
    else if (tok("dd")) { out += pad(d.getDate()); i += 2; }
    else if (tok("d")) { out += String(d.getDate()); i += 1; }
    else if (tok("HH")) { out += pad(d.getHours()); i += 2; }
    else if (tok("H")) { out += String(d.getHours()); i += 1; }
    else if (tok("hh")) { out += pad(h12); i += 2; }
    else if (tok("h")) { out += String(h12); i += 1; }
    else if (tok("mm")) { out += pad(d.getMinutes()); i += 2; }
    else if (tok("m")) { out += String(d.getMinutes()); i += 1; }
    else if (tok("ss")) { out += pad(d.getSeconds()); i += 2; }
    else if (tok("s")) { out += String(d.getSeconds()); i += 1; }
    else if (tok("AM/PM")) { out += d.getHours() < 12 ? "AM" : "PM"; i += 5; }
    else if (tok("am/pm")) { out += d.getHours() < 12 ? "am" : "pm"; i += 5; }
    else { out += fmt[i]; i += 1; }
  }
  return out;
}

/** Compare two literal operands; numeric when both parse as finite numbers, else string. */
function compare(a: string, op: IfOp, b: string): boolean {
  const na = Number(a);
  const nb = Number(b);
  const numeric = a.trim() !== "" && b.trim() !== "" && Number.isFinite(na) && Number.isFinite(nb);
  const [x, y]: [number | string, number | string] = numeric ? [na, nb] : [a, b];
  switch (op) {
    case "=": return x === y;
    case "<>": return x !== y;
    case "<": return x < y;
    case ">": return x > y;
    case "<=": return x <= y;
    case ">=": return x >= y;
  }
}

/** The chosen branch of an IF field, as fresh (cloned) runs. */
export function evaluateIf(spec: Extract<FieldSpec, { type: "IF" }>): Run[] {
  return cloneRuns(compare(spec.operandA, spec.op, spec.operandB) ? spec.trueRuns : spec.falseRuns);
}

/** Evaluate a built-in field to its display result. `baseStyle` styles DATE/TIME text. */
export function evaluateField(spec: FieldSpec, baseStyle: CharStyle, ctx: FieldEvalCtx): FieldEvalResult {
  switch (spec.type) {
    case "PAGE": return { kind: "token", token: `{page${numFmtToken(spec.numFmt)}}` };
    case "NUMPAGES": return { kind: "token", token: "{pages}" };
    case "DATE": return { kind: "runs", runs: [{ text: formatFieldDate(ctx.now, spec.format), style: { ...baseStyle } }] };
    case "TIME": return { kind: "runs", runs: [{ text: formatFieldDate(ctx.now, spec.format), style: { ...baseStyle } }] };
    case "IF": return { kind: "runs", runs: evaluateIf(spec) };
  }
}

// --- spec ⇄ instruction --------------------------------------------------------

const IF_OPS = new Set<string>(["=", "<>", "<", ">", "<=", ">="]);

function numFmtFromInstr(raw: string): PageNumFmt | undefined {
  if (/\\\*\s*ROMAN\b/.test(raw)) return "Roman";
  if (/\\\*\s*roman\b/.test(raw)) return "roman";
  if (/\\\*\s*ALPHABETIC\b/.test(raw)) return "Alpha";
  if (/\\\*\s*alphabetic\b/.test(raw)) return "alpha";
  return undefined;
}

/** The `\@ "..."` (or `\@ token`) date/time picture from an instruction. */
function atFormat(raw: string): string | undefined {
  return raw.match(/\\@\s*"([^"]*)"/)?.[1] ?? raw.match(/\\@\s*(\S+)/)?.[1];
}

/** Parse a built-in field instruction into a typed spec, or undefined (custom /
 *  unsupported / unparseable). IF result runs use `baseStyle` (plain text). */
export function parseFieldSpec(parsed: ParsedFieldInstruction, baseStyle: CharStyle): FieldSpec | undefined {
  switch (parsed.name) {
    case "PAGE": { const f = numFmtFromInstr(parsed.raw); return f ? { type: "PAGE", numFmt: f } : { type: "PAGE" }; }
    case "NUMPAGES": { const f = numFmtFromInstr(parsed.raw); return f ? { type: "NUMPAGES", numFmt: f } : { type: "NUMPAGES" }; }
    case "DATE": return { type: "DATE", format: atFormat(parsed.raw) ?? "M/d/yyyy" };
    case "TIME": return { type: "TIME", format: atFormat(parsed.raw) ?? "h:mm AM/PM" };
    case "IF": {
      const [operandA, op, operandB, trueText, falseText] = parsed.args;
      if (operandA === undefined || op === undefined || operandB === undefined || !IF_OPS.has(op)) return undefined;
      return {
        type: "IF", operandA, op: op as IfOp, operandB,
        trueRuns: [{ text: trueText ?? "", style: { ...baseStyle } }],
        falseRuns: [{ text: falseText ?? "", style: { ...baseStyle } }],
      };
    }
    default: return undefined;
  }
}

/** Synthesize the verbatim instruction for a typed spec (inverse of parseFieldSpec). */
export function buildInstruction(spec: FieldSpec): string {
  switch (spec.type) {
    case "PAGE": return ` PAGE${numFmtSwitch(spec.numFmt)} \\* MERGEFORMAT `;
    case "NUMPAGES": return ` NUMPAGES \\* MERGEFORMAT `;
    case "DATE": return ` DATE \\@ "${spec.format}" `;
    case "TIME": return ` TIME \\@ "${spec.format}" `;
    case "IF": return ` IF ${spec.operandA} ${spec.op} ${spec.operandB} "${textOfRuns(spec.trueRuns)}" "${textOfRuns(spec.falseRuns)}" `;
  }
}

/** Convenience: build a field name from a spec (uppercased keyword). */
export function fieldName(spec: FieldSpec): string {
  return spec.type;
}

export { parseFieldInstruction };
