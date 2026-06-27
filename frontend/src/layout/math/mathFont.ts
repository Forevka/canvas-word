// The metric/constant provider the math typesetter depends on. Abstracted behind
// an interface so the engine is PURE (testable with a stub) and so the real
// STIX Two Math OpenType MATH-table data can replace the defaults later (M4)
// without touching layoutMath.ts.
//
// `constants` are scaled per call by font size. The default ratios are the
// classic TeX values (a good approximation until the MATH table is wired in);
// real fonts carry their own in the MATH table, which this interface will expose.

import { fontMetrics, measureTextWidth } from "../metrics";
import { cloneFamilyFor, firstFamilyToken, MATH_FONT_FAMILY } from "../../fonts/clones";
import { STIX_MATH_CONSTANTS } from "./stixMathConstants";

export interface GlyphMetrics {
  width: number;
  ascent: number;
  descent: number;
  /** Italic correction — extra advance past the visual right edge of a slanted
   *  glyph, so an attached superscript clears it. */
  italicCorrection: number;
}

/** Layout constants in PIXELS for a given font size (TeX-style names). */
export interface MathConstants {
  axisHeight: number;
  ruleThickness: number;
  scriptPercent: number;
  scriptScriptPercent: number;
  fracNumShiftUp: number;
  fracDenShiftDown: number;
  fracNumGapMin: number;
  fracDenGapMin: number;
  supShiftUp: number;
  subShiftDown: number;
  supSubGapMin: number;
  spaceAfterScript: number;
  radicalRuleThickness: number;
  radicalVerticalGap: number;
  radicalExtraAscender: number;
  radicalDegreeRaise: number;
  /** Math spacing quanta (1mu = size/18): thin=3mu, med=4mu, thick=5mu. */
  thinSpace: number;
  medSpace: number;
  thickSpace: number;
  /** Minimum height of a display-style n-ary operator (∑ ∫). */
  displayOperatorMinHeight: number;
  /** Gaps/shifts for over/under limits on big operators. */
  upperLimitGapMin: number;
  lowerLimitGapMin: number;
  upperLimitBaselineRise: number;
  lowerLimitBaselineDrop: number;
}

export interface MathFont {
  family: string;
  /** Measure a glyph string at a size; `italic`/`bold` select the face. */
  measure(text: string, sizePx: number, italic: boolean, bold: boolean): GlyphMetrics;
  constants(sizePx: number): MathConstants;
}

/** TeX-derived default constants (placeholder for the STIX MATH table). */
export function defaultMathConstants(sizePx: number): MathConstants {
  const mu = sizePx / 18;
  const rule = Math.max(0.5, 0.04 * sizePx);
  return {
    axisHeight: 0.25 * sizePx,
    ruleThickness: rule,
    scriptPercent: 0.7,
    scriptScriptPercent: 0.5,
    fracNumShiftUp: 0.45 * sizePx,
    fracDenShiftDown: 0.55 * sizePx,
    fracNumGapMin: 3 * rule,
    fracDenGapMin: 3 * rule,
    supShiftUp: 0.42 * sizePx,
    subShiftDown: 0.2 * sizePx,
    supSubGapMin: 4 * rule,
    spaceAfterScript: 0.5 * mu,
    radicalRuleThickness: rule,
    radicalVerticalGap: 1.5 * rule,
    radicalExtraAscender: rule,
    radicalDegreeRaise: 0.6,
    thinSpace: 3 * mu,
    medSpace: 4 * mu,
    thickSpace: 5 * mu,
    displayOperatorMinHeight: 1.8 * sizePx,
    upperLimitGapMin: 0.2 * sizePx,
    lowerLimitGapMin: 0.2 * sizePx,
    upperLimitBaselineRise: 0.3 * sizePx,
    lowerLimitBaselineDrop: 0.6 * sizePx,
  };
}

/** Real STIX Two Math constants (extracted from its OpenType MATH table). Used
 *  for the bundled math font so positioning matches a professional math renderer;
 *  the TeX defaults above remain the fallback for the stub font in tests. */
export function stixMathConstants(sizePx: number): MathConstants {
  const s = STIX_MATH_CONSTANTS;
  const mu = sizePx / 18;
  return {
    axisHeight: s.axisHeight * sizePx,
    ruleThickness: s.ruleThickness * sizePx,
    scriptPercent: s.scriptPercent,
    scriptScriptPercent: s.scriptScriptPercent,
    fracNumShiftUp: s.fracNumShiftUp * sizePx,
    fracDenShiftDown: s.fracDenShiftDown * sizePx,
    fracNumGapMin: s.fracNumGapMin * sizePx,
    fracDenGapMin: s.fracDenGapMin * sizePx,
    supShiftUp: s.supShiftUp * sizePx,
    subShiftDown: s.subShiftDown * sizePx,
    supSubGapMin: s.supSubGapMin * sizePx,
    spaceAfterScript: s.spaceAfterScript * sizePx,
    radicalRuleThickness: s.radicalRuleThickness * sizePx,
    radicalVerticalGap: s.radicalVerticalGap * sizePx,
    radicalExtraAscender: s.radicalExtraAscender * sizePx,
    radicalDegreeRaise: s.radicalDegreeRaise,
    thinSpace: 3 * mu,
    medSpace: 4 * mu,
    thickSpace: 5 * mu,
    displayOperatorMinHeight: s.displayOperatorMinHeight * sizePx,
    upperLimitGapMin: s.upperLimitGapMin * sizePx,
    lowerLimitGapMin: s.lowerLimitGapMin * sizePx,
    upperLimitBaselineRise: s.upperLimitBaselineRiseMin * sizePx,
    lowerLimitBaselineDrop: s.lowerLimitBaselineDropMin * sizePx,
  };
}

/** Default provider backed by the existing clone-metric measurement surface.
 *  Uses the bundled metric clone of `family` so the editor, browser export, and
 *  Node backend measure identically. Per-glyph bounding boxes / real italic
 *  correction await the STIX MATH integration; until then ascent/descent come
 *  from the font's overall metrics and italic correction is a small estimate. */
export function createMathFont(family: string): MathFont {
  const clone = cloneFamilyFor(firstFamilyToken(family)).clone;
  const fontStr = (sizePx: number, italic: boolean, bold: boolean): string =>
    `${italic ? "italic " : ""}${bold ? "700" : "400"} ${sizePx}px ${clone}`;
  // The bundled math font carries real MATH-table constants; anything else
  // (a clone fallback) uses the TeX defaults.
  const constants = clone === MATH_FONT_FAMILY ? stixMathConstants : defaultMathConstants;
  return {
    family: clone,
    measure(text, sizePx, italic, bold) {
      const font = fontStr(sizePx, italic, bold);
      const width = measureTextWidth(text, font);
      const fm = fontMetrics(font);
      // Estimate italic correction from the slant when the glyph is italic.
      const italicCorrection = italic ? Math.min(0.12 * sizePx, 0.15 * width) : 0;
      return { width, ascent: fm.ascent, descent: fm.descent, italicCorrection };
    },
    constants,
  };
}
