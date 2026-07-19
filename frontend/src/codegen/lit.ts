// Serialize plain JSON-ish model values (style patches, borders, list levels,
// math ASTs) to VALID TypeScript source literals. Everything the reverse builder
// emits as a data argument passes through here, so it must round-trip a value to
// source that `eval`/`tsc` parse back to a structurally-equal value.
//
// Scope: strings, finite numbers, booleans, null, arrays, and plain objects with
// string keys — the shape of the document model's leaf data. `undefined` is kept
// (emitted as the identifier `undefined`) because a style patch uses it to STRIP
// an inherited field (e.g. `{ highlightColor: undefined }`); dropping it would
// change meaning. Functions, symbols, class instances, cyclic refs are out of
// scope (the model has none) and throw rather than emit garbage.

/** A bare-identifier object key needs no quoting (matches idiomatic hand-written
 *  builder code); anything else is emitted as a quoted string key. */
const IDENT = /^[A-Za-z_$][A-Za-z0-9_$]*$/;

/** Emit a JS string literal with a conservative escape set (double-quoted). */
function str(s: string): string {
  let out = '"';
  for (const ch of s) {
    const code = ch.codePointAt(0)!;
    if (ch === '"') out += '\\"';
    else if (ch === "\\") out += "\\\\";
    else if (ch === "\n") out += "\\n";
    else if (ch === "\r") out += "\\r";
    else if (ch === "\t") out += "\\t";
    else if (code < 0x20) out += "\\u" + code.toString(16).padStart(4, "0");
    else out += ch;
  }
  return out + '"';
}

/** Serialize `v` to a TS source literal. `indent` is the CURRENT indent string of
 *  the line the literal starts on; nested lines add two spaces. Objects/arrays
 *  print on one line when short, else break one entry per line for readability. */
export function lit(v: unknown, indent = ""): string {
  if (v === undefined) return "undefined";
  if (v === null) return "null";
  const t = typeof v;
  if (t === "string") return str(v as string);
  if (t === "boolean") return String(v);
  if (t === "number") {
    if (!Number.isFinite(v as number)) throw new Error(`cannot serialize non-finite number: ${String(v)}`);
    return String(v);
  }
  if (Array.isArray(v)) return arr(v, indent);
  if (t === "object") return obj(v as Record<string, unknown>, indent);
  throw new Error(`cannot serialize value of type ${t}`);
}

function arr(a: unknown[], indent: string): string {
  if (a.length === 0) return "[]";
  const inner = indent + "  ";
  const parts = a.map((x) => lit(x, inner));
  const oneLine = `[${parts.join(", ")}]`;
  if (oneLine.length <= 72 && !oneLine.includes("\n")) return oneLine;
  return `[\n${parts.map((p) => inner + p).join(",\n")},\n${indent}]`;
}

function obj(o: Record<string, unknown>, indent: string): string {
  const keys = Object.keys(o);
  if (keys.length === 0) return "{}";
  const inner = indent + "  ";
  const entries = keys.map((k) => `${IDENT.test(k) ? k : str(k)}: ${lit(o[k], inner)}`);
  const oneLine = `{ ${entries.join(", ")} }`;
  if (oneLine.length <= 72 && !oneLine.includes("\n")) return oneLine;
  return `{\n${entries.map((e) => inner + e).join(",\n")},\n${indent}}`;
}
