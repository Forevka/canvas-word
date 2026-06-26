// AST -> Presentation MathML string. Pure string building (no DOM), reusing the
// OOXML `el`/`escapeText` emitters — they are generic XML builders. The inverse
// of parse.ts; the two MUST round-trip for the supported element set.

import type { MathEquation, MathNode, MathVariant } from "@cw/shared";
import { el, escapeText } from "../export/docx/xmlWrite";

const MATHML_NS = "http://www.w3.org/1998/Math/MathML";

/** Serialize a whole equation to a standalone `<math>` string. */
export function serializeMathml(eq: MathEquation): string {
  return el(
    "math",
    { xmlns: MATHML_NS, display: eq.display ? "block" : "inline" },
    eq.root.children.map(nodeXml).join(""),
  );
}

/** Serialize a single node (no namespace) — used for fragments / previews. */
export function mathNodeToMathml(n: MathNode): string {
  return nodeXml(n);
}

function variantAttr(v: MathVariant | undefined): Record<string, string> | undefined {
  return v && v !== "italic" ? { mathvariant: v } : undefined;
}

function nodeXml(n: MathNode): string {
  switch (n.type) {
    case "row":
      return el("mrow", undefined, n.children.map(nodeXml).join(""));
    case "ident":
      return el("mi", variantAttr(n.variant), escapeText(n.text));
    case "number":
      return el("mn", undefined, escapeText(n.text));
    case "op": {
      const attrs: Record<string, string> = {};
      if (n.stretchy !== undefined) attrs.stretchy = String(n.stretchy);
      if (n.form) attrs.form = n.form;
      return el("mo", Object.keys(attrs).length ? attrs : undefined, escapeText(n.text));
    }
    case "text":
      return el("mtext", undefined, escapeText(n.text));
    case "space":
      return el("mspace", n.widthEm !== undefined ? { width: `${n.widthEm}em` } : undefined);
    case "frac": {
      const attrs: Record<string, string> = {};
      if (n.bevelled) attrs.bevelled = "true";
      if (n.thickness === "0") attrs.linethickness = "0";
      return el(
        "mfrac",
        Object.keys(attrs).length ? attrs : undefined,
        nodeXml(n.num) + nodeXml(n.den),
      );
    }
    case "script": {
      if (n.sub && n.sup)
        return el("msubsup", undefined, nodeXml(n.base) + nodeXml(n.sub) + nodeXml(n.sup));
      if (n.sup) return el("msup", undefined, nodeXml(n.base) + nodeXml(n.sup));
      if (n.sub) return el("msub", undefined, nodeXml(n.base) + nodeXml(n.sub));
      return nodeXml(n.base);
    }
    case "radical":
      return n.index
        ? el("mroot", undefined, nodeXml(n.radicand) + nodeXml(n.index))
        : el("msqrt", undefined, nodeXml(n.radicand));
    case "fenced": {
      const attrs: Record<string, string> = { open: n.open, close: n.close };
      if (n.separators !== undefined) attrs.separators = n.separators;
      return el("mfenced", attrs, nodeXml(n.child));
    }
    case "limit": {
      const accent = n.accent ? { accent: "true" } : undefined;
      if (n.under && n.over)
        return el("munderover", accent, nodeXml(n.base) + nodeXml(n.under) + nodeXml(n.over));
      if (n.over) return el("mover", accent, nodeXml(n.base) + nodeXml(n.over));
      if (n.under) return el("munder", accent, nodeXml(n.base) + nodeXml(n.under));
      return nodeXml(n.base);
    }
    case "nary": {
      // No MathML n-ary element: emit a scripted operator then the body.
      const opNode: MathNode = { type: "op", text: n.op, stretchy: false };
      const scripted: MathNode = n.hideOp
        ? { type: "row", children: [] }
        : { type: "script", base: opNode, ...(n.sub ? { sub: n.sub } : {}), ...(n.sup ? { sup: n.sup } : {}) };
      return el("mrow", undefined, nodeXml(scripted) + nodeXml(n.body));
    }
    case "matrix":
      return el(
        "mtable",
        undefined,
        n.rows
          .map((row) => el("mtr", undefined, row.map((c) => el("mtd", undefined, nodeXml(c))).join("")))
          .join(""),
      );
    case "phantom":
      return el("mphantom", undefined, nodeXml(n.child));
    case "unknown":
      // Re-emit verbatim MathML when we have it; otherwise a visible marker.
      return n.mathml ?? el("merror", undefined, el("mtext", undefined, "?"));
  }
}
