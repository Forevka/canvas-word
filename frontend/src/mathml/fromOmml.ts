// OMML (the `m:` namespace) -> AST. Runs at docx import on the already-parsed
// txml tree. The inverse of toOmml.ts.
//
// OMML doesn't distinguish mi/mn/mo — everything is an `m:r` with text and a
// style. We classify the run text into identifier / number / operator / text so
// the resulting MathML is semantically meaningful (spacing, italics).

import {
  emptyMathRow,
  type MathMatrix,
  type MathNode,
  type MathRow,
  type MathVariant,
} from "@cw/shared";
import { type TNode } from "txml/txml";
import { childByLocal, childEls, local, mathOn, mathVal } from "./xmlNode";

/** Convert an `m:oMath` / `m:oMathPara` element to an equation root row. */
export function ommlToMathml(oMath: TNode): MathRow {
  const el = local(oMath.tagName) === "oMathPara" ? childByLocal(oMath, "oMath") ?? oMath : oMath;
  return rowLike(childEls(el));
}

const OP_RE = /^[^\p{L}\p{N}\s]+$/u;

/** Property elements (`*Pr`) carry formatting, never content — skip in rows. */
const isProp = (n: TNode): boolean => local(n.tagName).endsWith("Pr");

/** Map content children to nodes, dropping property elements; result is a row. */
function rowLike(els: TNode[]): MathRow {
  return { type: "row", children: els.filter((e) => !isProp(e)).map(nodeFrom) };
}

/** A wrapper operand (`m:e`/`m:num`/`m:sub`/…) -> its content as a node. */
function slot(parent: TNode, name: string): MathNode {
  const c = childByLocal(parent, name);
  return c ? rowLike(childEls(c)) : emptyMathRow();
}

/** Raw (untrimmed) text of an `m:r`'s `m:t` children. */
function runText(r: TNode): string {
  let out = "";
  for (const t of childEls(r)) {
    if (local(t.tagName) !== "t") continue;
    for (const c of t.children) if (typeof c === "string") out += c;
  }
  return out;
}

function variantFromSty(sty: string | undefined): MathVariant | undefined {
  switch (sty) {
    case "b":
      return "bold";
    case "bi":
      return "bold-italic";
    case "p":
      return "normal";
    default:
      return undefined; // "i" / absent => MathML default (italic)
  }
}

function runNode(r: TNode): MathNode {
  const t = runText(r);
  const rPr = childByLocal(r, "rPr");
  const sty = rPr ? mathVal(rPr, "sty") : undefined;
  if (/^[0-9]+([.,][0-9]+)?$/.test(t)) return { type: "number", text: t };
  if (t.length > 0 && OP_RE.test(t)) return { type: "op", text: t };
  if (/\p{L}/u.test(t)) {
    const variant = variantFromSty(sty);
    return { type: "ident", text: t, ...(variant ? { variant } : {}) };
  }
  return { type: "text", text: t };
}

function nodeFrom(n: TNode): MathNode {
  switch (local(n.tagName)) {
    case "oMath":
    case "oMathPara":
    case "e":
    case "num":
    case "den":
    case "sub":
    case "sup":
    case "lim":
    case "deg":
    case "fName":
      return rowLike(childEls(n));
    case "r":
      return runNode(n);
    case "f": {
      const fPr = childByLocal(n, "fPr");
      const type = fPr ? mathVal(fPr, "type") : undefined;
      return {
        type: "frac",
        num: slot(n, "num"),
        den: slot(n, "den"),
        ...(type === "skw" ? { bevelled: true } : {}),
        ...(type === "noBar" || type === "lin" ? { thickness: "0" as const } : {}),
      };
    }
    case "sSup":
      return { type: "script", base: slot(n, "e"), sup: slot(n, "sup") };
    case "sSub":
      return { type: "script", base: slot(n, "e"), sub: slot(n, "sub") };
    case "sSubSup":
      return { type: "script", base: slot(n, "e"), sub: slot(n, "sub"), sup: slot(n, "sup") };
    case "sPre": {
      // Pre-scripts: best-effort approximation as post-scripts on the base.
      return { type: "script", base: slot(n, "e"), sub: slot(n, "sub"), sup: slot(n, "sup") };
    }
    case "rad": {
      const radPr = childByLocal(n, "radPr");
      const degHidden = radPr ? mathOn(radPr, "degHide") : false;
      const degEl = childByLocal(n, "deg");
      const hasIndex = !degHidden && degEl !== undefined && childEls(degEl).length > 0;
      return {
        type: "radical",
        radicand: slot(n, "e"),
        ...(hasIndex ? { index: slot(n, "deg") } : {}),
      };
    }
    case "d": {
      const dPr = childByLocal(n, "dPr");
      const open = dPr ? mathVal(dPr, "begChr") : undefined;
      const close = dPr ? mathVal(dPr, "endChr") : undefined;
      const sep = dPr ? mathVal(dPr, "sepChr") : undefined;
      const es = childEls(n).filter((c) => local(c.tagName) === "e");
      const child: MathNode =
        es.length === 1 ? rowLike(childEls(es[0]!)) : { type: "row", children: es.map((e) => rowLike(childEls(e))) };
      return {
        type: "fenced",
        open: open ?? "(",
        close: close ?? ")",
        ...(sep !== undefined ? { separators: sep } : {}),
        child,
      };
    }
    case "nary": {
      const naryPr = childByLocal(n, "naryPr");
      const op = (naryPr ? mathVal(naryPr, "chr") : undefined) ?? "∫";
      const subHidden = naryPr ? mathOn(naryPr, "subHide") : false;
      const supHidden = naryPr ? mathOn(naryPr, "supHide") : false;
      const subEl = childByLocal(n, "sub");
      const supEl = childByLocal(n, "sup");
      // An integral with no bounds has empty (or absent) m:sub/m:sup — omit them
      // rather than emitting empty limit slots.
      const sub = !subHidden && subEl && childEls(subEl).length > 0 ? rowLike(childEls(subEl)) : undefined;
      const sup = !supHidden && supEl && childEls(supEl).length > 0 ? rowLike(childEls(supEl)) : undefined;
      return {
        type: "nary",
        op,
        ...(sub ? { sub } : {}),
        ...(sup ? { sup } : {}),
        body: slot(n, "e"),
      };
    }
    case "limLow":
      return { type: "limit", base: slot(n, "e"), under: slot(n, "lim") };
    case "limUpp":
      return { type: "limit", base: slot(n, "e"), over: slot(n, "lim") };
    case "acc": {
      const accPr = childByLocal(n, "accPr");
      const chr = (accPr ? mathVal(accPr, "chr") : undefined) ?? "̂";
      return { type: "limit", base: slot(n, "e"), over: { type: "op", text: chr }, accent: true };
    }
    case "bar": {
      const barPr = childByLocal(n, "barPr");
      const pos = barPr ? mathVal(barPr, "pos") : "top";
      const bar: MathNode = { type: "op", text: pos === "bot" ? "_" : "‾", stretchy: true };
      return pos === "bot"
        ? { type: "limit", base: slot(n, "e"), under: bar }
        : { type: "limit", base: slot(n, "e"), over: bar };
    }
    case "groupChr": {
      const gPr = childByLocal(n, "groupChrPr");
      const chr = (gPr ? mathVal(gPr, "chr") : undefined) ?? "⏟";
      const pos = gPr ? mathVal(gPr, "pos") : "bot";
      const g: MathNode = { type: "op", text: chr, stretchy: true };
      return pos === "top"
        ? { type: "limit", base: slot(n, "e"), over: g }
        : { type: "limit", base: slot(n, "e"), under: g };
    }
    case "func":
      return { type: "row", children: [slot(n, "fName"), slot(n, "e")] };
    case "box":
    case "borderBox":
      return slot(n, "e");
    case "eqArr": {
      const rows = childEls(n)
        .filter((c) => local(c.tagName) === "e")
        .map((e): MathNode[] => [rowLike(childEls(e))]);
      return { type: "matrix", rows };
    }
    case "m":
      return matrixFrom(n);
    default:
      return childEls(n).length ? rowLike(childEls(n)) : { type: "unknown" };
  }
}

function matrixFrom(n: TNode): MathMatrix {
  const rows: MathNode[][] = [];
  for (const mr of childEls(n)) {
    if (local(mr.tagName) !== "mr") continue;
    rows.push(childEls(mr).filter((e) => local(e.tagName) === "e").map((e) => rowLike(childEls(e))));
  }
  return { type: "matrix", rows };
}
