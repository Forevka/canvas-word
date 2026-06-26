// MathML AST -> LaTeX. The (approximate) inverse of latex.ts, so the equation
// editor can convert between its LaTeX and MathML views in BOTH directions
// without losing the equation. Reuses the parser's symbol tables (inverted) as
// the single source of truth; anything unmapped falls back to its literal char,
// which the LaTeX parser reads straight back.

import type { MathNode, MathRow, MathVariant } from "@cw/shared";
import { ACCENTS, BIG, GREEK, OPS } from "./latex";

/** Invert a name→glyph table to glyph→`\name`, keeping the FIRST name as canonical
 *  (table order puts the preferred spelling first, e.g. `leq` before `le`). */
function invert(table: Record<string, string>): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, glyph] of Object.entries(table)) if (!(glyph in out)) out[glyph] = "\\" + name;
  return out;
}

const GREEK_R = invert(GREEK);
const OPS_R = invert(OPS);
const BIG_R = invert(BIG);
const ACCENT_R = invert(ACCENTS);

const VARIANT_CMD: Partial<Record<MathVariant, string>> = {
  "double-struck": "mathbb",
  bold: "mathbf",
  "bold-italic": "boldsymbol",
  normal: "mathrm",
  "sans-serif": "mathsf",
  monospace: "mathtt",
  script: "mathcal",
  fraktur: "mathfrak",
};

const ENV_FOR_FENCE: Record<string, string> = {
  "(": "pmatrix", "[": "bmatrix", "{": "Bmatrix", "|": "vmatrix", "‖": "Vmatrix",
};

/** Serialize a node; `tight` true means the result is a single atom (no braces
 *  needed when used as a script/argument). */
function node(n: MathNode): { tex: string; tight: boolean } {
  switch (n.type) {
    case "row": {
      const parts = n.children.map(node);
      return { tex: parts.map((p) => p.tex).join(" "), tight: parts.length <= 1 };
    }
    case "ident": {
      const base = GREEK_R[n.text] ?? n.text;
      if (n.variant && n.variant !== "italic" && VARIANT_CMD[n.variant]) {
        return { tex: `\\${VARIANT_CMD[n.variant]}{${base}}`, tight: true };
      }
      return { tex: base, tight: [...base].length === 1 };
    }
    case "number":
      return { tex: n.text, tight: [...n.text].length === 1 };
    case "op":
      return { tex: OPS_R[n.text] ?? BIG_R[n.text] ?? n.text, tight: [...n.text].length === 1 };
    case "text":
      return { tex: `\\text{${n.text}}`, tight: true };
    case "space":
      return { tex: spaceCmd(n.widthEm ?? 0), tight: true };
    case "frac":
      return {
        tex: (n.thickness === "0" ? "\\binom" : "\\frac") + `{${inner(n.num)}}{${inner(n.den)}}`,
        tight: true,
      };
    case "script": {
      const base = arg(n.base);
      const sub = n.sub !== undefined ? `_{${inner(n.sub)}}` : "";
      const sup = n.sup !== undefined ? `^{${inner(n.sup)}}` : "";
      return { tex: `${base}${sub}${sup}`, tight: false };
    }
    case "radical":
      return {
        tex: n.index ? `\\sqrt[${inner(n.index)}]{${inner(n.radicand)}}` : `\\sqrt{${inner(n.radicand)}}`,
        tight: true,
      };
    case "fenced":
      return { tex: fenced(n), tight: true };
    case "limit": {
      // Accent (hat/bar/vec…) → the accent command.
      if (n.accent && n.over && (n.over.type === "op" || n.over.type === "ident")) {
        const cmd = ACCENT_R[n.over.text];
        if (cmd) return { tex: `${cmd}{${inner(n.base)}}`, tight: true };
      }
      // Big operator with limits → the natural ∑_{}^{} form (round-trips, and
      // the layout puts the scripts under/over in display style).
      if (n.base.type === "op" && n.base.text in BIG_R) {
        const sub = n.under !== undefined ? `_{${inner(n.under)}}` : "";
        const sup = n.over !== undefined ? `^{${inner(n.over)}}` : "";
        return { tex: `${BIG_R[n.base.text]}${sub}${sup}`, tight: false };
      }
      let t = arg(n.base);
      if (n.under) t = `\\underset{${inner(n.under)}}{${t}}`;
      if (n.over) t = `\\overset{${inner(n.over)}}{${t}}`;
      return { tex: t, tight: true };
    }
    case "nary": {
      const op = BIG_R[n.op] ?? OPS_R[n.op] ?? n.op;
      const sub = n.sub !== undefined ? `_{${inner(n.sub)}}` : "";
      const sup = n.sup !== undefined ? `^{${inner(n.sup)}}` : "";
      return { tex: `${op}${sub}${sup} ${inner(n.body)}`, tight: false };
    }
    case "matrix":
      return { tex: matrix(n.rows, "matrix"), tight: true };
    case "phantom":
      return { tex: `\\phantom{${inner(n.child)}}`, tight: true };
    case "unknown":
      return { tex: n.mathml ? "" : "", tight: true };
  }
}

/** Bare (unbraced) serialization of a node's contents. */
function inner(n: MathNode): string {
  return node(n).tex;
}

/** A node usable as a script base / argument — braced when it isn't a single atom. */
function arg(n: MathNode): string {
  const { tex, tight } = node(n);
  return tight ? tex : `{${tex}}`;
}

function spaceCmd(em: number): string {
  if (em >= 2) return "\\qquad";
  if (em >= 1) return "\\quad";
  if (em >= 0.27) return "\\;";
  if (em >= 0.2) return "\\:";
  if (em > 0) return "\\,";
  if (em < 0) return "\\!";
  return "";
}

function fenced(n: Extract<MathNode, { type: "fenced" }>): string {
  // A delimited matrix becomes \begin{pmatrix} … (nicer than \left( … \right)).
  if (n.child.type === "matrix") {
    const env = ENV_FOR_FENCE[n.open];
    if (env) return matrix(n.child.rows, env);
  }
  const open = n.open || ".";
  const close = n.close || ".";
  return `\\left${delim(open)} ${inner(n.child)} \\right${delim(close)}`;
}

function delim(ch: string): string {
  if (ch === "{" || ch === "}") return "\\" + ch;
  if (ch === "‖") return "\\|";
  if (ch === "⟨") return "\\langle";
  if (ch === "⟩") return "\\rangle";
  if (ch === "⌊") return "\\lfloor";
  if (ch === "⌋") return "\\rfloor";
  if (ch === "⌈") return "\\lceil";
  if (ch === "⌉") return "\\rceil";
  return ch;
}

function matrix(rows: MathNode[][], env: string): string {
  const body = rows.map((r) => r.map(inner).join(" & ")).join(" \\\\ ");
  return `\\begin{${env}} ${body} \\end{${env}}`;
}

/** Serialize an equation root to a LaTeX string. */
export function mathToLatex(root: MathRow): string {
  return root.children.map(inner).join(" ").replace(/\s+/g, " ").trim();
}
