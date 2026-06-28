// Equation input adapters for the builder: turn an author-friendly LaTeX or
// MathML string into the model's MathEquation AST. Both converters are the same
// ones the editor uses (parse.ts / latex.ts), so a builder equation is identical
// to one typed in the visual editor — and round-trips to OMML on docx export.

import type { MathEquation } from "@cw/shared";
import { latexToMath } from "../mathml/latex";
import { parseMathml } from "../mathml/parse";

/** A MathEquation from a LaTeX source (e.g. `\frac{a}{b}`). `display` picks block
 *  (centered, own line) vs inline rendering. */
export function equationFromLatex(latex: string, display: boolean): MathEquation {
  return { root: latexToMath(latex), display };
}

/** A MathEquation from a presentation-MathML string (`<math>…</math>` or a bare
 *  fragment). The string's own `display` attribute is overridden by `display`. */
export function equationFromMathml(mathml: string, display: boolean): MathEquation {
  return { ...parseMathml(mathml), display };
}
