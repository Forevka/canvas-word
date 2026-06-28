// LaTeX (math mode) -> MathML AST. A focused recursive-descent parser over the
// common TeX math subset, so users can type `\frac{-b\pm\sqrt{b^2-4ac}}{2a}`
// instead of MathML. Lenient: unknown commands become literal text rather than
// errors, so a half-typed formula still previews.
//
// Covers: fractions (\frac \dfrac \tfrac \binom), radicals (\sqrt[n]{}), scripts
// (^ _), fences (\left..\right and bare brackets), matrices/cases (\begin{..}),
// accents (\hat \bar \vec ..), big operators with limits (\sum \int ..), Greek,
// the common operator/relation symbols, function names, \mathbb/\mathbf/.. styles,
// \text, and spacing. Not a full LaTeX engine — enough for authoring equations.

import type { MathNary, MathNode, MathRow, MathVariant } from "@cw/shared";
import { emptyMathRow } from "@cw/shared";

// ── Symbol tables ────────────────────────────────────────────────────────────

export const GREEK: Record<string, string> = {
  alpha: "α", beta: "β", gamma: "γ", delta: "δ", epsilon: "ε", varepsilon: "ε",
  zeta: "ζ", eta: "η", theta: "θ", vartheta: "ϑ", iota: "ι", kappa: "κ",
  lambda: "λ", mu: "μ", nu: "ν", xi: "ξ", omicron: "ο", pi: "π", varpi: "ϖ",
  rho: "ρ", varrho: "ϱ", sigma: "σ", varsigma: "ς", tau: "τ", upsilon: "υ",
  phi: "φ", varphi: "ϕ", chi: "χ", psi: "ψ", omega: "ω",
  Gamma: "Γ", Delta: "Δ", Theta: "Θ", Lambda: "Λ", Xi: "Ξ", Pi: "Π",
  Sigma: "Σ", Upsilon: "Υ", Phi: "Φ", Psi: "Ψ", Omega: "Ω",
};

// Commands that become a single operator/relation/fence glyph (an `mo`).
export const OPS: Record<string, string> = {
  times: "×", div: "÷", pm: "±", mp: "∓", cdot: "⋅", ast: "∗", star: "⋆",
  leq: "≤", le: "≤", geq: "≥", ge: "≥", neq: "≠", ne: "≠", approx: "≈",
  equiv: "≡", cong: "≅", sim: "∼", simeq: "≃", propto: "∝", asymp: "≍",
  subset: "⊂", supset: "⊃", subseteq: "⊆", supseteq: "⊇", in: "∈", notin: "∉",
  ni: "∋", cup: "∪", cap: "∩", vee: "∨", wedge: "∧", oplus: "⊕", otimes: "⊗",
  to: "→", rightarrow: "→", leftarrow: "←", gets: "←", leftrightarrow: "↔",
  Rightarrow: "⇒", Leftarrow: "⇐", Leftrightarrow: "⇔", implies: "⟹", iff: "⟺",
  mapsto: "↦", uparrow: "↑", downarrow: "↓", infty: "∞", partial: "∂",
  nabla: "∇", forall: "∀", exists: "∃", nexists: "∄", emptyset: "∅",
  varnothing: "∅", angle: "∠", perp: "⊥", parallel: "∥", mid: "∣",
  cdots: "⋯", ldots: "…", dots: "…", vdots: "⋮", ddots: "⋱", prime: "′",
  circ: "∘", bullet: "∙", setminus: "∖", smallsetminus: "∖", oslash: "⊘",
  langle: "⟨", rangle: "⟩", lfloor: "⌊", rfloor: "⌋", lceil: "⌈", rceil: "⌉",
  backslash: "\\", land: "∧", lor: "∨", neg: "¬", lnot: "¬", top: "⊤", bot: "⊥",
  Re: "ℜ", Im: "ℑ", aleph: "ℵ", hbar: "ℏ", ell: "ℓ", wp: "℘", Box: "□",
  dagger: "†", ddagger: "‡", surd: "√", triangle: "△", square: "□",
  cong2: "≅", doteq: "≐", models: "⊨", vdash: "⊢", dashv: "⊣", because: "∵",
  therefore: "∴", colon: ":", semicolon: ";", lbrace: "{", rbrace: "}",
};

// Big operators (∑ ∫ …) — rendered as `mo`; limits handled by the layout engine.
export const BIG: Record<string, string> = {
  sum: "∑", prod: "∏", coprod: "∐", int: "∫", iint: "∬", iiint: "∭",
  oint: "∮", bigcup: "⋃", bigcap: "⋂", bigvee: "⋁", bigwedge: "⋀",
  bigoplus: "⨁", bigotimes: "⨂", bigodot: "⨀", biguplus: "⨄", bigsqcup: "⨆",
};

// Function names — multi-letter identifiers, rendered upright.
const FUNCS = new Set([
  "sin", "cos", "tan", "cot", "sec", "csc", "sinh", "cosh", "tanh", "coth",
  "arcsin", "arccos", "arctan", "log", "ln", "lg", "exp", "lim", "limsup",
  "liminf", "max", "min", "sup", "inf", "det", "gcd", "deg", "dim", "ker",
  "hom", "arg", "Pr", "mod",
]);

// \mathXX style commands → mathvariant.
const STYLES: Record<string, MathVariant> = {
  mathbb: "double-struck", mathbf: "bold", boldsymbol: "bold", mathit: "italic",
  mathrm: "normal", mathsf: "sans-serif", mathtt: "monospace",
  mathcal: "script", mathscr: "script", mathfrak: "fraktur",
};

// Accent commands → the over-glyph placed on the base.
export const ACCENTS: Record<string, string> = {
  hat: "^", widehat: "^", bar: "‾", overline: "‾", vec: "→", overrightarrow: "→",
  tilde: "~", widetilde: "~", dot: "˙", ddot: "¨", check: "ˇ", breve: "˘",
  acute: "´", grave: "`", mathring: "˚",
};

const FENCES: Record<string, string> = {
  "(": "(", ")": ")", "[": "[", "]": "]", "\\{": "{", "\\}": "}",
  "\\langle": "⟨", "\\rangle": "⟩", "\\lfloor": "⌊", "\\rfloor": "⌋",
  "\\lceil": "⌈", "\\rceil": "⌉", "|": "|", "\\|": "‖", ".": "",
  "\\vert": "|", "\\Vert": "‖", "\\lvert": "|", "\\rvert": "|",
};

const ENV_FENCES: Record<string, [string, string]> = {
  matrix: ["", ""], pmatrix: ["(", ")"], bmatrix: ["[", "]"],
  Bmatrix: ["{", "}"], vmatrix: ["|", "|"], Vmatrix: ["‖", "‖"],
  cases: ["{", ""], smallmatrix: ["", ""],
};

// ── Tokenizer ────────────────────────────────────────────────────────────────

type Tok =
  | { k: "cmd"; v: string }
  | { k: "char"; v: string }
  | { k: "rawtext"; v: string }
  | { k: "{" }
  | { k: "}" }
  | { k: "^" }
  | { k: "_" }
  | { k: "&" }
  | { k: "\\\\" };

// Commands whose `{...}` argument is TEXT, not math: whitespace inside is
// significant (`\text{hello world}`). The tokenizer captures that brace group
// verbatim as a single `rawtext` token instead of math-tokenizing it (which
// would drop the spaces). Nested braces are kept as literal characters.
const RAW_TEXT_CMDS = new Set(["text", "mbox", "operatorname"]);

/** Capture a `{...}` group starting at `src[start]` ("{") as its inner text,
 *  whitespace preserved, balancing nested braces. An ESCAPED brace (`\{` / `\}`)
 *  is copied verbatim and never counted toward nesting — that pairs with the
 *  escaping done by toLatex so `\text{a\_b}` / `\text{a\}b}` round-trip. Returns
 *  the inner string and the index just past the closing brace. If `src[start]`
 *  isn't "{", returns null. */
function captureBraceText(src: string, start: number): { text: string; end: number } | null {
  if (src[start] !== "{") return null;
  let depth = 0;
  let buf = "";
  for (let k = start; k < src.length; k++) {
    const c = src[k]!;
    if (c === "\\" && k + 1 < src.length) {
      buf += c + src[k + 1]!; // escaped char — copy both, skip brace counting
      k++;
      continue;
    }
    if (c === "{") {
      depth++;
      if (depth === 1) continue; // skip the outermost opening brace
    } else if (c === "}") {
      depth--;
      if (depth === 0) return { text: buf, end: k + 1 };
    }
    buf += c;
  }
  return { text: buf, end: src.length }; // unbalanced — take the rest
}

/** Inverse of toLatex's `\text{…}` escaping. Single-pass so a decoded char is
 *  never re-scanned. Longest control words first to avoid prefix ambiguity. */
const TEXT_UNESCAPES: [string, string][] = [
  ["\\textbackslash{}", "\\"],
  ["\\textasciitilde{}", "~"],
  ["\\textasciicircum{}", "^"],
  ["\\{", "{"],
  ["\\}", "}"],
  ["\\$", "$"],
  ["\\&", "&"],
  ["\\#", "#"],
  ["\\_", "_"],
  ["\\%", "%"],
];

function unescapeTextLatex(s: string): string {
  let out = "";
  let i = 0;
  outer: while (i < s.length) {
    if (s[i] === "\\") {
      for (const [from, to] of TEXT_UNESCAPES) {
        if (s.startsWith(from, i)) {
          out += to;
          i += from.length;
          continue outer;
        }
      }
    }
    out += s[i]!;
    i++;
  }
  return out;
}

function tokenize(src: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  while (i < src.length) {
    const ch = src[i]!;
    if (ch === "\\") {
      // Control word (\alpha) or control symbol (\{, \\, \,).
      const rest = src.slice(i + 1);
      const word = /^[A-Za-z]+/.exec(rest);
      if (word) {
        out.push({ k: "cmd", v: word[0] });
        i += 1 + word[0].length;
        // \text{…}-family: grab the following brace group as raw (spaced) text.
        if (RAW_TEXT_CMDS.has(word[0])) {
          let j = i;
          while (j < src.length && /\s/.test(src[j]!)) j++; // skip space before {
          const grp = captureBraceText(src, j);
          if (grp) {
            out.push({ k: "rawtext", v: unescapeTextLatex(grp.text) });
            i = grp.end;
          }
        }
      } else if (rest[0] === "\\") {
        out.push({ k: "\\\\" });
        i += 2;
      } else {
        out.push({ k: "cmd", v: rest[0] ?? "" }); // \{ \} \, \; \! \space
        i += 2;
      }
      continue;
    }
    if (ch === "{") { out.push({ k: "{" }); i++; continue; }
    if (ch === "}") { out.push({ k: "}" }); i++; continue; }
    if (ch === "^") { out.push({ k: "^" }); i++; continue; }
    if (ch === "_") { out.push({ k: "_" }); i++; continue; }
    if (ch === "&") { out.push({ k: "&" }); i++; continue; }
    if (/\s/.test(ch)) { i++; continue; } // whitespace is not significant in math
    out.push({ k: "char", v: ch });
    i++;
  }
  return out;
}

// ── Parser ───────────────────────────────────────────────────────────────────

class Parser {
  private i = 0;
  constructor(private toks: Tok[]) {}

  private peek(): Tok | undefined {
    return this.toks[this.i];
  }
  private next(): Tok | undefined {
    return this.toks[this.i++];
  }

  /** Parse a list of atoms until a stop predicate (or end), applying scripts. */
  parseList(stop?: (t: Tok) => boolean): MathNode[] {
    const out: MathNode[] = [];
    for (;;) {
      const t = this.peek();
      if (!t || (stop && stop(t))) break;
      if (t.k === "^" || t.k === "_") {
        this.applyScript(out, t.k);
        continue;
      }
      const node = this.parseAtom();
      if (node) out.push(node);
    }
    return out;
  }

  private applyScript(out: MathNode[], kind: "^" | "_"): void {
    this.next(); // consume ^ or _
    const arg = this.parseAtom() ?? emptyMathRow();
    const base = out.pop() ?? emptyMathRow();
    if (base.type === "script" || base.type === "nary") {
      // x^a_b style, or \sum_a^b — fill the missing bound slot in place. For an
      // n-ary, this is what gives \sum/\int real OMML limits instead of scripts.
      if (kind === "^") base.sup = arg;
      else base.sub = arg;
      out.push(base);
    } else {
      out.push(kind === "^" ? { type: "script", base, sup: arg } : { type: "script", base, sub: arg });
    }
  }

  /** Parse a big operator (∑ ∫ ∏ …) into a real n-ary node: consume its bounds
   *  (`_`/`^`, in either order) and then the operand it applies to.
   *
   *  The operand binds as the immediately following FACTOR (an atom plus any of
   *  its own scripts), which lands in `nary.body` so OMML export fills the
   *  `<m:e>` base instead of letting the summand float outside the `<m:nary>`.
   *  `\sum_{i=1}^{n} i`, `\int_a^b f`, and `\sum_i x^2` all bind correctly.
   *
   *  known limitation: in LaTeX a big operator scopes over the rest of the
   *  expression up to the next lower-precedence boundary; we bind only the next
   *  factor. Multi-term or parenthesised summands (`\sum_i a_i b_i`,
   *  `\int (x+1) dx`) keep just the first factor in the body — the rest stay as
   *  siblings. Full precedence handling is out of scope for this fix. */
  private naryOp(op: string): MathNode {
    const nary: MathNary = { type: "nary", op, body: emptyMathRow() };
    while (this.peek()?.k === "_" || this.peek()?.k === "^") {
      const kind = (this.next() as { k: "_" | "^" }).k;
      const bound = this.parseAtom() ?? emptyMathRow();
      if (kind === "^") nary.sup = bound;
      else nary.sub = bound;
    }
    const next = this.peek();
    if (next && startsNaryBody(next)) {
      nary.body = this.parseFactor() ?? nary.body;
    }
    return nary;
  }

  /** Parse one factor — an atom together with any scripts applied to it
   *  (`x`, `x^2`, `a_i^2`). Reuses the same script-folding as `parseList`. */
  private parseFactor(): MathNode | null {
    const atom = this.parseAtom();
    if (atom === null) return null;
    const stack: MathNode[] = [atom];
    while (this.peek()?.k === "^" || this.peek()?.k === "_") {
      this.applyScript(stack, (this.peek() as { k: "^" | "_" }).k);
    }
    return stack[stack.length - 1] ?? null;
  }

  /** Parse a single atom (group, command, or character). */
  private parseAtom(): MathNode | null {
    const t = this.next();
    if (!t) return null;
    switch (t.k) {
      case "{": {
        const list = this.parseList((x) => x.k === "}");
        if (this.peek()?.k === "}") this.next(); // consume closing brace
        return this.row(list, true);
      }
      case "}":
        return null; // unbalanced — ignore
      case "&":
      case "\\\\":
        return null; // only meaningful inside environments
      case "char":
        return charNode(t.v);
      case "rawtext":
        return { type: "text", text: t.v }; // stray text group (no preceding \text)
      case "cmd":
        return this.command(t.v);
      default:
        return null;
    }
  }

  private row(children: MathNode[], collapse = false): MathNode {
    if (collapse && children.length === 1) return children[0]!;
    return { type: "row", children };
  }

  /** Consume one mandatory `{...}` argument (or the next single atom). */
  private arg(): MathNode {
    if (this.peek()?.k === "{") {
      this.next();
      const list = this.parseList((x) => x.k === "}");
      if (this.peek()?.k === "}") this.next(); // consume closing brace
      return this.row(list, true);
    }
    return this.parseAtom() ?? emptyMathRow();
  }

  /** Consume a `\text`-family argument as raw (whitespace-preserving) text. The
   *  tokenizer already captured the brace group as one `rawtext` token; fall back
   *  to a math atom flattened to text if it wasn't (e.g. `\text x`). */
  private rawTextArg(): string {
    if (this.peek()?.k === "rawtext") return (this.next() as { v: string }).v;
    return plainText(this.arg());
  }

  /** Consume an optional `[...]` argument; returns null if absent. */
  private optArg(): MathNode | null {
    if (this.peek()?.k === "char" && (this.peek() as { v: string }).v === "[") {
      this.next();
      const inner = this.parseList((x) => x.k === "char" && (x as { v: string }).v === "]");
      if (this.peek()?.k === "char" && (this.peek() as { v: string }).v === "]") this.next();
      return this.row(inner, true);
    }
    return null;
  }

  private command(name: string): MathNode | null {
    if (name in GREEK) return { type: "ident", text: GREEK[name]! };
    if (name in OPS) return { type: "op", text: OPS[name]! };
    // Big operators are true n-ary objects: `_`/`^` fill their bounds and the
    // following operand (summand / integrand) binds into the body, so OMML
    // export emits a real `<m:nary>` with a populated `<m:e>` base.
    if (name in BIG) return this.naryOp(BIG[name]!);
    if (FUNCS.has(name)) return { type: "ident", text: name, variant: "normal" };
    if (name in STYLES) {
      const variant = STYLES[name]!;
      const inner = this.arg();
      return applyVariant(inner, variant);
    }
    if (name in ACCENTS) {
      return { type: "limit", base: this.arg(), over: { type: "op", text: ACCENTS[name]! }, accent: true };
    }
    switch (name) {
      case "frac":
      case "dfrac":
      case "tfrac":
        return { type: "frac", num: this.arg(), den: this.arg() };
      case "binom":
      case "choose":
        return { type: "fenced", open: "(", close: ")", child: { type: "frac", num: this.arg(), den: this.arg(), thickness: "0" } };
      case "sqrt": {
        const index = this.optArg();
        const radicand = this.arg();
        return index ? { type: "radical", radicand, index } : { type: "radical", radicand };
      }
      case "text":
      case "mbox":
      case "operatorname":
        return { type: "text", text: this.rawTextArg() };
      case "overset":
      case "stackrel": {
        const over = this.arg();
        return { type: "limit", base: this.arg(), over };
      }
      case "underset": {
        const under = this.arg();
        return { type: "limit", base: this.arg(), under };
      }
      case "phantom":
        // Only full \phantom maps cleanly: the model has no horizontal/vertical
        // axis, so \hphantom / \vphantom are left to fall through (they would
        // change meaning if round-tripped back as a plain \phantom).
        return { type: "phantom", child: this.arg() };
      case "left":
        return this.leftRight();
      case "right":
        return null; // handled by leftRight
      case "begin":
        return this.environment();
      case "end":
        return null;
      case "quad":
        return { type: "space", widthEm: 1 };
      case "qquad":
        return { type: "space", widthEm: 2 };
      case ",":
        return { type: "space", widthEm: 0.167 };
      case ":":
      case ">":
        return { type: "space", widthEm: 0.222 };
      case ";":
        return { type: "space", widthEm: 0.278 };
      case "!":
        return { type: "space", widthEm: -0.167 };
      case " ":
        return { type: "space", widthEm: 0.25 };
      case "{":
        return { type: "op", text: "{" };
      case "}":
        return { type: "op", text: "}" };
      case "|":
        return { type: "op", text: "‖" };
      case "lim":
        return { type: "ident", text: "lim", variant: "normal" };
      default:
        // Unknown command — show its name as text rather than failing.
        return { type: "text", text: name };
    }
  }

  /** \left<delim> … \right<delim>. */
  private leftRight(): MathNode {
    const open = this.delim();
    const inner = this.parseList((t) => t.k === "cmd" && t.v === "right");
    if (this.peek()?.k === "cmd" && (this.peek() as { v: string }).v === "right") this.next();
    const close = this.delim();
    return { type: "fenced", open, close, child: this.row(inner) };
  }

  /** Read a delimiter token after \left / \right (a char like ( or a command). */
  private delim(): string {
    const t = this.next();
    if (!t) return "";
    if (t.k === "char") return FENCES[t.v] ?? t.v;
    if (t.k === "cmd") return FENCES["\\" + t.v] ?? OPS[t.v] ?? "";
    return "";
  }

  /** \begin{env} … \end{env} — matrices and cases. */
  private environment(): MathNode {
    const envName = plainText(this.arg());
    const fence = ENV_FENCES[envName] ?? ["", ""];
    const rows: MathNode[][] = [];
    let row: MathNode[] = [];
    let cell: MathNode[] = [];
    const flushCell = (): void => { row.push(this.row(cell)); cell = []; };
    const flushRow = (): void => { flushCell(); rows.push(row); row = []; };
    for (;;) {
      const t = this.peek();
      if (!t) break;
      if (t.k === "cmd" && t.v === "end") { this.next(); this.arg(); break; }
      if (t.k === "&") { this.next(); flushCell(); continue; }
      if (t.k === "\\\\") { this.next(); flushRow(); continue; }
      const node = this.parseAtom();
      if (node) cell.push(node);
    }
    if (cell.length > 0 || row.length > 0) flushRow();
    const matrix: MathNode = { type: "matrix", rows };
    if (!fence[0] && !fence[1]) return matrix;
    return { type: "fenced", open: fence[0], close: fence[1], child: matrix };
  }
}

// ── Helpers ──────────────────────────────────────────────────────────────────

const OP_CHARS = new Set("+-*/=<>(),.;:!|[]'".split("").concat(["±", "×", "÷"]));

/** Whether `t` begins an operand that a big operator should bind as its body.
 *  Groups, commands, and `\text` runs do; so do alphanumeric atoms. Operators,
 *  relations, punctuation, fences, and the structural tokens (`^ _ } & \\`) do
 *  NOT — they parse as ordinary siblings, leaving the body empty (e.g. the
 *  trailing `= S` in `\sum_{i} a_i = S`). */
function startsNaryBody(t: Tok): boolean {
  switch (t.k) {
    case "{":
    case "rawtext":
    case "cmd":
      return true;
    case "char":
      return /[A-Za-z0-9]/.test(t.v);
    default:
      return false;
  }
}

function charNode(ch: string): MathNode {
  if (/[0-9]/.test(ch)) return { type: "number", text: ch };
  if (/[A-Za-z]/.test(ch)) return { type: "ident", text: ch };
  if (OP_CHARS.has(ch)) return { type: "op", text: ch };
  return { type: "op", text: ch };
}

/** Re-tag identifiers/numbers inside a node with a math variant (for `\mathbb{R}`,
 *  `\mathbb{1}`, `\mathtt{0}`, …). Numbers carry the style too so blackboard /
 *  monospace digits survive (mirrors the identifier path). */
function applyVariant(node: MathNode, variant: MathVariant): MathNode {
  if (node.type === "ident") return { ...node, variant };
  if (node.type === "number") return { ...node, variant };
  if (node.type === "row") return { ...node, children: node.children.map((c) => applyVariant(c, variant)) };
  return node;
}

/** Flatten a node to plain text (for \text / env names). */
function plainText(node: MathNode): string {
  switch (node.type) {
    case "ident":
    case "number":
    case "op":
    case "text":
      return node.text;
    case "row":
      return node.children.map(plainText).join("");
    default:
      return "";
  }
}

/** Parse a LaTeX math string into an equation root row. */
export function latexToMath(src: string): MathRow {
  const parser = new Parser(tokenize(src));
  return { type: "row", children: parser.parseList() };
}
