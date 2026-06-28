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

/** The single child of a one-element row (else undefined) — used to peek through
 *  the row wrapper that `slot` always produces. */
function soleChild(n: MathNode): MathNode | undefined {
  return n.type === "row" && n.children.length === 1 ? n.children[0] : undefined;
}

/** Interleave separator glyphs between delimiter (`m:d`) operands. `sep` is the
 *  `m:sepChr` value: `undefined` defaults to `,`; an empty string means no
 *  separator; multiple chars cycle (as MathML `mfenced` separators do). */
function interleaveSeparators(operands: MathNode[], sep: string | undefined): MathNode[] {
  const chars = sep ?? ",";
  if (chars.length === 0) return operands;
  const out: MathNode[] = [];
  operands.forEach((operand, i) => {
    if (i > 0) out.push({ type: "op", text: chars[(i - 1) % chars.length]! });
    out.push(operand);
  });
  return out;
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

/** OMML run props (`m:sty` + `m:scr`) -> MathML variant. Inverse of
 *  toOmml.runPropsFor: the SCRIPT family (`m:scr`) selects the alphanumeric
 *  style, with `m:sty` adding bold/italic. */
function variantFrom(sty: string | undefined, scr: string | undefined): MathVariant | undefined {
  const bold = sty === "b" || sty === "bi";
  switch (scr) {
    case "double-struck":
      return "double-struck";
    case "fraktur":
      return bold ? "bold-fraktur" : "fraktur";
    case "script":
      return bold ? "bold-script" : "script";
    case "sans-serif":
      if (sty === "bi") return "sans-serif-bold-italic";
      if (sty === "i") return "sans-serif-italic";
      if (sty === "b") return "bold-sans-serif";
      return "sans-serif";
    case "monospace":
      return "monospace";
    default:
      return variantFromSty(sty);
  }
}

function runNode(r: TNode): MathNode {
  const t = runText(r);
  const rPr = childByLocal(r, "rPr");
  const sty = rPr ? mathVal(rPr, "sty") : undefined;
  const scr = rPr ? mathVal(rPr, "scr") : undefined;
  // A whitespace-only run is intentional spacing (we have no other mspace
  // representation in OMML) — recover it as a `space` node, ~0.25em per space.
  if (t.length > 0 && /^\s+$/.test(t)) return { type: "space", widthEm: t.length * 0.25 };
  if (/^[0-9]+([.,][0-9]+)?$/.test(t)) return { type: "number", text: t };
  if (t.length > 0 && OP_RE.test(t)) return { type: "op", text: t };
  if (/\p{L}/u.test(t)) {
    const variant = variantFrom(sty, scr);
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
      const operands = childEls(n)
        .filter((c) => local(c.tagName) === "e")
        .map((e) => rowLike(childEls(e)));
      // A multi-operand delimiter (e.g. the tuple `(a,b)`) carries its separator
      // in `m:sepChr` (default `,`). The layout/LaTeX paths don't read
      // `fenced.separators`, so materialize the separators as real `op` nodes
      // between operands instead of storing them in that unused field.
      const child: MathNode =
        operands.length <= 1
          ? operands[0] ?? emptyMathRow()
          : { type: "row", children: interleaveSeparators(operands, sep) };
      return {
        type: "fenced",
        open: open ?? "(",
        close: close ?? ")",
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
    case "limLow": {
      // Combined under+over (munderover) exports as nested limUpp/limLow; collapse
      // the inner partner back into a single limit so it round-trips.
      const base = slot(n, "e");
      const under = slot(n, "lim");
      const inner = soleChild(base);
      if (inner && inner.type === "limit" && inner.over && !inner.under && !inner.accent)
        return { type: "limit", base: inner.base, under, over: inner.over };
      return { type: "limit", base, under };
    }
    case "limUpp": {
      const base = slot(n, "e");
      const over = slot(n, "lim");
      const inner = soleChild(base);
      if (inner && inner.type === "limit" && inner.under && !inner.over && !inner.accent)
        return { type: "limit", base: inner.base, under: inner.under, over };
      return { type: "limit", base, over };
    }
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
