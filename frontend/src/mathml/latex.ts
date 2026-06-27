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

import type { MathNode, MathRow, MathVariant } from "@cw/shared";
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
  | { k: "{" }
  | { k: "}" }
  | { k: "^" }
  | { k: "_" }
  | { k: "&" }
  | { k: "\\\\" };

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
    if (base.type === "script") {
      // x^a_b style — fill the missing slot.
      if (kind === "^") base.sup = arg;
      else base.sub = arg;
      out.push(base);
    } else {
      out.push(kind === "^" ? { type: "script", base, sup: arg } : { type: "script", base, sub: arg });
    }
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
    if (name in BIG) return { type: "op", text: BIG[name]! };
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
        return { type: "text", text: plainText(this.arg()) };
      case "overset":
      case "stackrel": {
        const over = this.arg();
        return { type: "limit", base: this.arg(), over };
      }
      case "underset": {
        const under = this.arg();
        return { type: "limit", base: this.arg(), under };
      }
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

function charNode(ch: string): MathNode {
  if (/[0-9]/.test(ch)) return { type: "number", text: ch };
  if (/[A-Za-z]/.test(ch)) return { type: "ident", text: ch };
  if (OP_CHARS.has(ch)) return { type: "op", text: ch };
  return { type: "op", text: ch };
}

/** Re-tag identifiers inside a node with a math variant (for \mathbb{R} etc.). */
function applyVariant(node: MathNode, variant: MathVariant): MathNode {
  if (node.type === "ident") return { ...node, variant };
  if (node.type === "number" && variant === "bold") return node;
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
