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

type Sty = "i" | "b" | "bi" | "p";
type Scr = "double-struck" | "script" | "fraktur" | "sans-serif" | "monospace";

/** OMML run properties for a MathML variant: `m:sty` (bold/italic) plus an
 *  optional `m:scr` (the alphanumeric SCRIPT family — blackboard / script /
 *  fraktur / sans-serif / monospace). OMML's math-run default is italic;
 *  numbers/operators/text are plain ("p"). The inverse lives in fromOmml. */
function runPropsFor(variant: MathVariant | undefined): { sty: Sty; scr?: Scr } {
  switch (variant) {
    case undefined:
    case "italic":
      return { sty: "i" };
    case "bold":
      return { sty: "b" };
    case "bold-italic":
      return { sty: "bi" };
    case "normal":
      return { sty: "p" };
    case "double-struck":
      return { sty: "p", scr: "double-struck" };
    case "script":
      return { sty: "p", scr: "script" };
    case "bold-script":
      return { sty: "b", scr: "script" };
    case "fraktur":
      return { sty: "p", scr: "fraktur" };
    case "bold-fraktur":
      return { sty: "b", scr: "fraktur" };
    case "sans-serif":
      return { sty: "p", scr: "sans-serif" };
    case "bold-sans-serif":
      return { sty: "b", scr: "sans-serif" };
    case "sans-serif-italic":
      return { sty: "i", scr: "sans-serif" };
    case "sans-serif-bold-italic":
      return { sty: "bi", scr: "sans-serif" };
    case "monospace":
      return { sty: "p", scr: "monospace" };
  }
}

const needsPreserve = (s: string): boolean => /^\s|\s$|\s\s/.test(s) || s.includes("\t");

/** A math run: `m:r` with an optional `m:rPr` (`m:scr` then `m:sty`, schema
 *  order) and the `m:t` text. */
function runXml(content: string, sty: Sty, scr?: Scr): string {
  const scrEl = scr ? el("m:scr", { "m:val": scr }) : "";
  const styEl = sty === "i" ? "" : el("m:sty", { "m:val": sty });
  const inner = scrEl + styEl;
  const rPr = inner ? el("m:rPr", undefined, inner) : "";
  const t = el("m:t", needsPreserve(content) ? { "xml:space": "preserve" } : undefined, escapeText(content));
  return el("m:r", undefined, rPr + t);
}

/** Approximate an `mspace` width (ems) as a run of literal spaces — OMML has no
 *  `mspace`. ~0.25em per space; negative/zero widths emit nothing. */
function spaceRunXml(widthEm: number | undefined): string {
  const em = widthEm ?? 0.25;
  if (em <= 0) return "";
  const n = Math.max(1, Math.round(em / 0.25));
  return runXml(" ".repeat(n), "p");
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
    case "ident": {
      const { sty, scr } = runPropsFor(n.variant);
      return runXml(n.text, sty, scr);
    }
    case "number":
      return runXml(n.text, "p");
    case "op":
    case "text":
      return runXml(n.text, "p");
    case "space":
      return spaceRunXml(n.widthEm);
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
