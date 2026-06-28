// Layer 1 (math): the canonical internal form for equations — a Presentation
// MathML abstract syntax tree. Pure data (JSON-serializable, no DOM/canvas), so
// it lives beside the rest of the document model and round-trips through
// persistence for free.
//
// WHY MathML and not OMML: the editor, caret movement, and the typesetter all
// operate on this structured tree. `.docx` stores math as OMML (the `m:`
// namespace), so the import/export pipelines convert OMML <-> this AST at the
// edges (see frontend/src/mathml). MathML strings (paste, the equation editor)
// also parse to / serialize from this same tree.
//
// Design notes for the editor:
//  - Script kinds (sup/sub/subsup) are ONE node (`script`) with optional sub/sup.
//  - Limit kinds (under/over/underover) are ONE node (`limit`).
//  - Radical kinds (sqrt/root) are ONE node (`radical`) with optional index.
//    Unifying these keeps the visual editor generic: it edits SLOTS, not a
//    per-element zoo. `mathSlots`/`mapMathSlots` expose those slots uniformly.
//  - An empty `row` (no children) is a PLACEHOLDER SLOT — what the editor shows
//    as a dotted box you can tab into and type.

/** MathML `mathvariant` — selects an alphanumeric style (blackboard, fraktur,
 *  script, …). Maps to/from OMML run style (`m:sty`) for the common cases. */
export type MathVariant =
  | "normal"
  | "bold"
  | "italic"
  | "bold-italic"
  | "double-struck"
  | "fraktur"
  | "bold-fraktur"
  | "script"
  | "bold-script"
  | "sans-serif"
  | "bold-sans-serif"
  | "sans-serif-italic"
  | "sans-serif-bold-italic"
  | "monospace";

/** A horizontal list of sub-expressions (MathML `mrow`; also the canonical
 *  wrapper for `math`/`mstyle`). An EMPTY row is an editable placeholder slot. */
export interface MathRow {
  type: "row";
  children: MathNode[];
}

/** Identifier — a variable or function name (MathML `mi`). Default rendering is
 *  italic for single letters; `variant` overrides (e.g. blackboard-bold ℝ). */
export interface MathIdent {
  type: "ident";
  text: string;
  variant?: MathVariant;
}

/** Number literal (MathML `mn`). Rendered upright by default; `variant` overrides
 *  the alphanumeric style (e.g. blackboard-bold `𝟙`, monospace digits). */
export interface MathNumber {
  type: "number";
  text: string;
  variant?: MathVariant;
}

/** Operator, relation, fence, or separator (MathML `mo`). `stretchy` lets a
 *  delimiter/arrow grow to its context; `form` disambiguates spacing class. */
export interface MathOperator {
  type: "op";
  text: string;
  stretchy?: boolean;
  form?: "prefix" | "infix" | "postfix";
}

/** Literal text run inside math (MathML `mtext`) — upright, normal spacing. */
export interface MathText {
  type: "text";
  text: string;
}

/** Explicit spacing (MathML `mspace`), width in ems. */
export interface MathSpace {
  type: "space";
  widthEm?: number;
}

/** Fraction (MathML `mfrac`). `bevelled` = diagonal bar; `thickness: "0"` = no
 *  bar (binomial coefficients `\binom`). */
export interface MathFrac {
  type: "frac";
  num: MathNode;
  den: MathNode;
  bevelled?: boolean;
  thickness?: "0" | "normal";
}

/** Sub/superscript (MathML `msub`/`msup`/`msubsup`), unified: presence of
 *  `sub`/`sup` selects the concrete element. At least one is set. */
export interface MathScript {
  type: "script";
  base: MathNode;
  sub?: MathNode;
  sup?: MathNode;
}

/** Radical (MathML `msqrt`/`mroot`). `index` present = nth root. */
export interface MathRadical {
  type: "radical";
  radicand: MathNode;
  index?: MathNode;
}

/** Fenced expression — delimiters around content (MathML `mfenced`, or OMML
 *  `m:d`). Empty `open`/`close` strings mean "no delimiter on that side". */
export interface MathFenced {
  type: "fenced";
  open: string;
  close: string;
  /** Separator chars cycled between children when the content is a row. */
  separators?: string;
  child: MathNode;
}

/** Under/over scripts (MathML `munder`/`mover`/`munderover`), unified — limits,
 *  overbars, hats, underbraces. `accent` keeps the OVER script tight to the base
 *  (MathML `accent`); `accentUnder` does the same for the UNDER script (MathML
 *  `accentunder`). The two are independent so an over-accent (hat) and an
 *  under-accent (underbrace/underbar) never collapse to the same encoding. */
export interface MathLimit {
  type: "limit";
  base: MathNode;
  under?: MathNode;
  over?: MathNode;
  accent?: boolean;
  accentUnder?: boolean;
}

/** N-ary operator with limits (OMML `m:nary`: ∑ ∏ ∫ …). MathML has no single
 *  element for this; it serializes to a scripted operator followed by the body.
 *  `hideOp` renders just the scripted limits with no operator glyph. */
export interface MathNary {
  type: "nary";
  op: string;
  sub?: MathNode;
  sup?: MathNode;
  body: MathNode;
  hideOp?: boolean;
}

/** Matrix / array / cases (MathML `mtable`, OMML `m:m`). Row-major. */
export interface MathMatrix {
  type: "matrix";
  rows: MathNode[][];
  colAlign?: ("left" | "center" | "right")[];
}

/** Phantom — occupies space but paints nothing (MathML `mphantom`). */
export interface MathPhantom {
  type: "phantom";
  child: MathNode;
}

/** Integrity net for constructs outside the supported set: the original source
 *  is kept verbatim so export re-emits it losslessly and nothing is silently
 *  dropped. The typesetter renders a visible placeholder. Reaching this in
 *  practice is a coverage bug to fix, not an accepted fallback. */
export interface MathUnknown {
  type: "unknown";
  omml?: string;
  mathml?: string;
}

export type MathNode =
  | MathRow
  | MathIdent
  | MathNumber
  | MathOperator
  | MathText
  | MathSpace
  | MathFrac
  | MathScript
  | MathRadical
  | MathFenced
  | MathLimit
  | MathNary
  | MathMatrix
  | MathPhantom
  | MathUnknown;

/** A whole equation as stored in the document: the root is always a `row`.
 *  `display` = block (own line, centered); inline = sits in a text line. */
export interface MathEquation {
  root: MathRow;
  display: boolean;
}

// ── Constructors / guards ────────────────────────────────────────────────────

/** An empty row — the editor's placeholder slot. */
export function emptyMathRow(): MathRow {
  return { type: "row", children: [] };
}

/** A row is a placeholder when it has no children (renders as a dotted box). */
export function isEmptyMathRow(n: MathNode): boolean {
  return n.type === "row" && n.children.length === 0;
}

/** Wrap a node in a row unless it already is one (normalizes slot contents). */
export function asMathRow(n: MathNode): MathRow {
  return n.type === "row" ? n : { type: "row", children: [n] };
}

// ── Generic slot traversal (powers the visual editor + the typesetter) ───────

/** Ordered child SLOTS of a node — the editable sub-expressions, in document
 *  order. Leaf nodes (ident/number/op/text/space/unknown) have none. Used by
 *  the editor for caret navigation and by converters/typesetter to recurse
 *  without a per-type switch at every call site. */
export function mathSlots(n: MathNode): MathNode[] {
  switch (n.type) {
    case "row":
      return n.children;
    case "frac":
      return [n.num, n.den];
    case "script":
      return [n.base, ...(n.sub ? [n.sub] : []), ...(n.sup ? [n.sup] : [])];
    case "radical":
      return [n.radicand, ...(n.index ? [n.index] : [])];
    case "fenced":
      return [n.child];
    case "limit":
      return [n.base, ...(n.under ? [n.under] : []), ...(n.over ? [n.over] : [])];
    case "nary":
      return [...(n.sub ? [n.sub] : []), ...(n.sup ? [n.sup] : []), n.body];
    case "matrix":
      return n.rows.flat();
    case "phantom":
      return [n.child];
    default:
      return [];
  }
}

/** Rebuild `n` with every slot replaced by `f(slot)` — an IMMUTABLE structural
 *  map (never mutates the input). The traversal order matches `mathSlots`. */
export function mapMathSlots(n: MathNode, f: (child: MathNode) => MathNode): MathNode {
  switch (n.type) {
    case "row":
      return { ...n, children: n.children.map(f) };
    case "frac":
      return { ...n, num: f(n.num), den: f(n.den) };
    case "script":
      return {
        ...n,
        base: f(n.base),
        ...(n.sub ? { sub: f(n.sub) } : {}),
        ...(n.sup ? { sup: f(n.sup) } : {}),
      };
    case "radical":
      return { ...n, radicand: f(n.radicand), ...(n.index ? { index: f(n.index) } : {}) };
    case "fenced":
      return { ...n, child: f(n.child) };
    case "limit":
      return {
        ...n,
        base: f(n.base),
        ...(n.under ? { under: f(n.under) } : {}),
        ...(n.over ? { over: f(n.over) } : {}),
      };
    case "nary":
      return {
        ...n,
        ...(n.sub ? { sub: f(n.sub) } : {}),
        ...(n.sup ? { sup: f(n.sup) } : {}),
        body: f(n.body),
      };
    case "matrix":
      return { ...n, rows: n.rows.map((r) => r.map(f)) };
    case "phantom":
      return { ...n, child: f(n.child) };
    default:
      return n;
  }
}

/** Depth-first visit of every node (self first, then slots). */
export function walkMath(n: MathNode, visit: (node: MathNode) => void): void {
  visit(n);
  for (const c of mathSlots(n)) walkMath(c, visit);
}

/** Plain-text approximation of an equation (for search, a11y, alt text). */
export function mathToPlainText(n: MathNode): string {
  switch (n.type) {
    case "ident":
    case "number":
    case "op":
    case "text":
      return n.text;
    case "space":
      return " ";
    case "frac":
      return `${mathToPlainText(n.num)}/${mathToPlainText(n.den)}`;
    case "fenced":
      return `${n.open}${mathToPlainText(n.child)}${n.close}`;
    case "nary":
      return n.op + mathSlots(n).map(mathToPlainText).join("");
    default:
      return mathSlots(n).map(mathToPlainText).join("");
  }
}
