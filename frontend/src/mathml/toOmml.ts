// AST -> OMML (Office Math Markup Language, the `m:` namespace) string. This is
// what `.docx` stores; emitted at the docx-export sites. Pure string building on
// the OOXML `el`/`escapeText` emitters. The inverse of fromOmml.ts.
//
// The `m:` namespace must be declared on the document root (see WML_NS in
// export/docx/xmlWrite.ts) for the produced markup to be valid.

import type { MathEquation, MathNode, MathVariant } from "@cw/shared";
import { el, escapeText } from "../export/docx/xmlWrite";

/** Serialize an equation to OMML. Inline math -> `m:oMath`; display math ->
 *  `m:oMathPara` wrapping it (Word's block-equation container). */
export function mathmlToOmml(eq: MathEquation): string {
  const oMath = el("m:oMath", undefined, eq.root.children.map(nodeXml).join(""));
  return eq.display ? el("m:oMathPara", undefined, oMath) : oMath;
}

/** OMML run style (`m:sty`) for a variant. OMML's math-run default is italic;
 *  numbers/operators/text are plain ("p"). */
function styFor(variant: MathVariant | undefined): "i" | "b" | "bi" | "p" {
  switch (variant) {
    case undefined:
    case "italic":
      return "i";
    case "bold":
      return "b";
    case "bold-italic":
      return "bi";
    default:
      return "p";
  }
}

const needsPreserve = (s: string): boolean => /^\s|\s$|\s\s/.test(s) || s.includes("\t");

/** A math run: `m:r` with an optional `m:rPr/m:sty` and the `m:t` text. */
function runXml(content: string, sty: "i" | "b" | "bi" | "p"): string {
  const rPr = sty === "i" ? "" : el("m:rPr", undefined, el("m:sty", { "m:val": sty }));
  const t = el("m:t", needsPreserve(content) ? { "xml:space": "preserve" } : undefined, escapeText(content));
  return el("m:r", undefined, rPr + t);
}

/** `m:e`/`m:num`/… operand wrapper holding a node's serialized content. */
const wrap = (tag: string, n: MathNode): string => el(tag, undefined, nodeXml(n));

/** First textual char of a node (for accent chars). */
function firstChar(n: MathNode): string {
  switch (n.type) {
    case "op":
    case "ident":
    case "text":
    case "number":
      return n.text;
    case "row":
      return n.children[0] ? firstChar(n.children[0]) : "";
    default:
      return "";
  }
}

function nodeXml(n: MathNode): string {
  switch (n.type) {
    case "row":
      return n.children.map(nodeXml).join("");
    case "ident":
      return runXml(n.text, styFor(n.variant));
    case "number":
      return runXml(n.text, "p");
    case "op":
    case "text":
      return runXml(n.text, "p");
    case "space":
      return runXml(" ", "p");
    case "frac": {
      const type = n.thickness === "0" ? "noBar" : n.bevelled ? "skw" : undefined;
      const fPr = type ? el("m:fPr", undefined, el("m:type", { "m:val": type })) : "";
      return el("m:f", undefined, fPr + wrap("m:num", n.num) + wrap("m:den", n.den));
    }
    case "script": {
      if (n.sub && n.sup)
        return el("m:sSubSup", undefined, wrap("m:e", n.base) + wrap("m:sub", n.sub) + wrap("m:sup", n.sup));
      if (n.sup) return el("m:sSup", undefined, wrap("m:e", n.base) + wrap("m:sup", n.sup));
      return el("m:sSub", undefined, wrap("m:e", n.base) + wrap("m:sub", n.sub!));
    }
    case "radical": {
      const radPr = n.index ? "" : el("m:radPr", undefined, el("m:degHide", { "m:val": "1" }));
      const deg = n.index ? wrap("m:deg", n.index) : el("m:deg");
      return el("m:rad", undefined, radPr + deg + wrap("m:e", n.radicand));
    }
    case "fenced": {
      const dPr = el(
        "m:dPr",
        undefined,
        el("m:begChr", { "m:val": n.open }) +
          el("m:endChr", { "m:val": n.close }) +
          (n.separators !== undefined ? el("m:sepChr", { "m:val": n.separators }) : ""),
      );
      return el("m:d", undefined, dPr + wrap("m:e", n.child));
    }
    case "limit": {
      if (n.over && n.accent && !n.under) {
        const accPr = el("m:accPr", undefined, el("m:chr", { "m:val": firstChar(n.over) || "̂" }));
        return el("m:acc", undefined, accPr + wrap("m:e", n.base));
      }
      const lower = n.under
        ? el("m:limLow", undefined, wrap("m:e", n.base) + wrap("m:lim", n.under))
        : undefined;
      if (n.over) {
        const e = lower ? el("m:e", undefined, lower) : wrap("m:e", n.base);
        return el("m:limUpp", undefined, e + wrap("m:lim", n.over));
      }
      return lower ?? wrap("m:e", n.base);
    }
    case "nary": {
      const naryPr = el(
        "m:naryPr",
        undefined,
        el("m:chr", { "m:val": n.op }) +
          (n.sub ? "" : el("m:subHide", { "m:val": "1" })) +
          (n.sup ? "" : el("m:supHide", { "m:val": "1" })),
      );
      const sub = n.sub ? wrap("m:sub", n.sub) : el("m:sub");
      const sup = n.sup ? wrap("m:sup", n.sup) : el("m:sup");
      return el("m:nary", undefined, naryPr + sub + sup + wrap("m:e", n.body));
    }
    case "matrix":
      return el(
        "m:m",
        undefined,
        n.rows
          .map((row) => el("m:mr", undefined, row.map((c) => wrap("m:e", c)).join("")))
          .join(""),
      );
    case "phantom":
      return nodeXml(n.child);
    case "unknown":
      return n.omml ?? runXml("?", "p");
  }
}
