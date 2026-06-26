// Presentation MathML string -> AST. The inverse of serialize.ts. Uses txml
// (worker-safe, ordered tree) like the docx importer. Namespace prefixes are
// matched by LOCAL name, so `<math>`, `<m:math>`, and `<mml:math>` all parse.

import { isElementNode, parse, type TNode } from "txml/txml";
import {
  emptyMathRow,
  type MathEquation,
  type MathMatrix,
  type MathNode,
  type MathRow,
  type MathVariant,
} from "@cw/shared";
import { attr, childEls, local, text } from "./xmlNode";

const KNOWN_VARIANTS = new Set<string>([
  "normal", "bold", "italic", "bold-italic", "double-struck", "fraktur",
  "bold-fraktur", "script", "bold-script", "sans-serif", "bold-sans-serif",
  "sans-serif-italic", "sans-serif-bold-italic", "monospace",
]);

/** Map a list of element children to nodes, collapsing to a single node or a row. */
const nodesOf = (els: TNode[]): MathNode[] => els.map(nodeFrom);
const rowOf = (els: TNode[]): MathRow => ({ type: "row", children: nodesOf(els) });
const slot = (els: TNode[], i: number): MathNode => (els[i] ? nodeFrom(els[i]!) : emptyMathRow());
/** A content slot holding ALL children: a lone child stays itself (no redundant
 *  mrow), several become a row, none becomes an empty placeholder. Keeping this
 *  collapsed is what makes parse∘serialize a fixed point. */
const collapse = (els: TNode[]): MathNode =>
  els.length === 0 ? emptyMathRow() : els.length === 1 ? nodeFrom(els[0]!) : rowOf(els);

/** Parse a standalone MathML string into an equation. Tolerant: a bare fragment
 *  (no `<math>` wrapper) is accepted and wrapped. */
export function parseMathml(xml: string): MathEquation {
  let roots: TNode[];
  try {
    roots = parse(xml, { keepWhitespace: true, decodeEntities: true }).filter(isElementNode);
  } catch {
    return { root: emptyMathRow(), display: false };
  }
  const mathEl = findMath(roots);
  if (!mathEl) {
    // No <math> wrapper — treat the top-level elements as the row contents.
    return { root: rowOf(roots), display: false };
  }
  return {
    root: rowOf(childEls(mathEl)),
    display: (attr(mathEl, "display") ?? attr(mathEl, "mode")) === "block",
  };
}

function findMath(roots: TNode[]): TNode | undefined {
  for (const r of roots) {
    if (local(r.tagName) === "math") return r;
    const deep = childEls(r).find((c) => local(c.tagName) === "math");
    if (deep) return deep;
  }
  return undefined;
}

function nodeFrom(n: TNode): MathNode {
  const els = childEls(n);
  switch (local(n.tagName)) {
    case "math":
    case "mrow":
    case "mstyle":
    case "mpadded":
      return rowOf(els);
    case "semantics":
      // The presentation node is the FIRST child; the rest are <annotation> /
      // <annotation-xml> metadata, not renderable content.
      return els[0] ? nodeFrom(els[0]!) : emptyMathRow();
    case "mi": {
      const v = attr(n, "mathvariant");
      return { type: "ident", text: text(n), ...(v && KNOWN_VARIANTS.has(v) ? { variant: v as MathVariant } : {}) };
    }
    case "mn":
      return { type: "number", text: text(n) };
    case "mo": {
      const stretchy = attr(n, "stretchy");
      const form = attr(n, "form");
      return {
        type: "op",
        text: text(n),
        ...(stretchy !== undefined ? { stretchy: stretchy !== "false" } : {}),
        ...(form === "prefix" || form === "infix" || form === "postfix" ? { form } : {}),
      };
    }
    case "mtext":
      return { type: "text", text: text(n) };
    case "mspace": {
      const w = parseEm(attr(n, "width"));
      return { type: "space", ...(w !== undefined ? { widthEm: w } : {}) };
    }
    case "mfrac": {
      const bevelled = attr(n, "bevelled") === "true";
      const noBar = attr(n, "linethickness") === "0" || attr(n, "linethickness") === "0pt";
      return {
        type: "frac",
        num: slot(els, 0),
        den: slot(els, 1),
        ...(bevelled ? { bevelled: true } : {}),
        ...(noBar ? { thickness: "0" } : {}),
      };
    }
    case "msup":
      return { type: "script", base: slot(els, 0), sup: slot(els, 1) };
    case "msub":
      return { type: "script", base: slot(els, 0), sub: slot(els, 1) };
    case "msubsup":
      return { type: "script", base: slot(els, 0), sub: slot(els, 1), sup: slot(els, 2) };
    case "msqrt":
      return { type: "radical", radicand: collapse(els) };
    case "mroot":
      return { type: "radical", radicand: slot(els, 0), index: slot(els, 1) };
    case "mfenced":
      return {
        type: "fenced",
        open: attr(n, "open") ?? "(",
        close: attr(n, "close") ?? ")",
        ...(attr(n, "separators") !== undefined ? { separators: attr(n, "separators")! } : {}),
        child: collapse(els),
      };
    case "mover":
      return { type: "limit", base: slot(els, 0), over: slot(els, 1), ...accentFlag(n) };
    case "munder":
      return { type: "limit", base: slot(els, 0), under: slot(els, 1), ...accentFlag(n) };
    case "munderover":
      return { type: "limit", base: slot(els, 0), under: slot(els, 1), over: slot(els, 2), ...accentFlag(n) };
    case "mtable":
      return matrixFrom(els);
    case "mphantom":
      return { type: "phantom", child: collapse(els) };
    default:
      // Unknown element: keep children editable rather than dropping anything.
      return els.length ? rowOf(els) : { type: "unknown" };
  }
}

function accentFlag(n: TNode): { accent?: true } {
  return attr(n, "accent") === "true" ? { accent: true } : {};
}

function matrixFrom(rowEls: TNode[]): MathMatrix {
  const rows: MathNode[][] = [];
  for (const tr of rowEls) {
    if (local(tr.tagName) !== "mtr") continue;
    rows.push(childEls(tr).filter((td) => local(td.tagName) === "mtd").map((td) => collapse(childEls(td))));
  }
  return { type: "matrix", rows };
}

function parseEm(w: string | undefined): number | undefined {
  if (!w) return undefined;
  const m = /^(-?(?:\d+(?:\.\d+)?|\.\d+))\s*em$/.exec(w.trim());
  if (!m) return undefined;
  const n = Number(m[1]);
  return Number.isFinite(n) ? n : undefined;
}
