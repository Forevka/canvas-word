// Build-time extraction of the OpenType MATH table's MathConstants from the
// bundled STIX Two Math font. fontkit does NOT parse the MATH table, so we read
// the raw SFNT bytes ourselves and emit per-em fractions into a TS data module
// the runtime consumes (keeping the runtime zero-dependency, no font parsing).
//
//   node scripts/extract-math-constants.mjs
//
// Re-run if the math font is replaced. Output: src/layout/math/stixMathConstants.ts

import { readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

const FONT = fileURLToPath(new URL("../src/export/shared/fonts/StixTwoMath-Regular.ttf", import.meta.url));
const OUT = fileURLToPath(new URL("../src/layout/math/stixMathConstants.ts", import.meta.url));

const buf = readFileSync(FONT);
const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
const u16 = (o) => dv.getUint16(o);
const i16 = (o) => dv.getInt16(o);
const u32 = (o) => dv.getUint32(o);

// --- SFNT table directory: find 'MATH' and 'head' (unitsPerEm) ---------------
const numTables = u16(4);
let mathOff = 0;
let headOff = 0;
for (let i = 0; i < numTables; i++) {
  const rec = 12 + i * 16;
  const tag = String.fromCharCode(buf[rec], buf[rec + 1], buf[rec + 2], buf[rec + 3]);
  const off = u32(rec + 8);
  if (tag === "MATH") mathOff = off;
  if (tag === "head") headOff = off;
}
if (!mathOff) throw new Error("no MATH table in font");
if (!headOff) throw new Error("no head table in font (cannot read unitsPerEm)");
const unitsPerEm = u16(headOff + 18);

// --- MATH header -> MathConstants subtable -----------------------------------
const constantsOff = mathOff + u16(mathOff + 4); // mathConstantsOffset

// Walk the MathConstants table in spec order. Two field widths:
//   plain  = a single int16/uint16 (percentages + the two min-heights)
//   record = MathValueRecord (int16 value + uint16 device offset) -> take value
let p = constantsOff;
const plain = () => { const v = i16(p); p += 2; return v; };
const record = () => { const v = i16(p); p += 4; return v; };

const scriptPercentScaleDown = plain();
const scriptScriptPercentScaleDown = plain();
const delimitedSubFormulaMinHeight = u16(p); p += 2;
const displayOperatorMinHeight = u16(p); p += 2;
const mathLeading = record();
const axisHeight = record();
const accentBaseHeight = record();
const flattenedAccentBaseHeight = record();
const subscriptShiftDown = record();
const subscriptTopMax = record();
const subscriptBaselineDropMin = record();
const superscriptShiftUp = record();
const superscriptShiftUpCramped = record();
const superscriptBottomMin = record();
const superscriptBaselineDropMax = record();
const subSuperscriptGapMin = record();
const superscriptBottomMaxWithSubscript = record();
const spaceAfterScript = record();
const upperLimitGapMin = record();
const upperLimitBaselineRiseMin = record();
const lowerLimitGapMin = record();
const lowerLimitBaselineDropMin = record();
const stackTopShiftUp = record();
const stackTopDisplayStyleShiftUp = record();
const stackBottomShiftDown = record();
const stackBottomDisplayStyleShiftDown = record();
const stackGapMin = record();
const stackDisplayStyleGapMin = record();
const stretchStackTopShiftUp = record();
const stretchStackBottomShiftDown = record();
const stretchStackGapAboveMin = record();
const stretchStackGapBelowMin = record();
const fractionNumeratorShiftUp = record();
const fractionNumeratorDisplayStyleShiftUp = record();
const fractionDenominatorShiftDown = record();
const fractionDenominatorDisplayStyleShiftDown = record();
const fractionNumeratorGapMin = record();
const fractionNumDisplayStyleGapMin = record();
const fractionRuleThickness = record();
const fractionDenominatorGapMin = record();
const fractionDenomDisplayStyleGapMin = record();
const skewedFractionHorizontalGap = record();
const skewedFractionVerticalGap = record();
const overbarVerticalGap = record();
const overbarRuleThickness = record();
const overbarExtraAscender = record();
const underbarVerticalGap = record();
const underbarRuleThickness = record();
const underbarExtraDescender = record();
const radicalVerticalGap = record();
const radicalDisplayStyleVerticalGap = record();
const radicalRuleThickness = record();
const radicalExtraAscender = record();
const radicalKernBeforeDegree = record();
const radicalKernAfterDegree = record();
const radicalDegreeBottomRaisePercent = plain();

const em = (u) => Math.round((u / unitsPerEm) * 100000) / 100000;

// Per-em fractions consumed by mathFont.ts (multiplied by font size at runtime).
const data = {
  axisHeight: em(axisHeight),
  ruleThickness: em(fractionRuleThickness),
  scriptPercent: scriptPercentScaleDown / 100,
  scriptScriptPercent: scriptScriptPercentScaleDown / 100,
  fracNumShiftUp: em(fractionNumeratorShiftUp),
  fracDenShiftDown: em(fractionDenominatorShiftDown),
  fracNumGapMin: em(fractionNumeratorGapMin),
  fracDenGapMin: em(fractionDenominatorGapMin),
  supShiftUp: em(superscriptShiftUp),
  supShiftUpCramped: em(superscriptShiftUpCramped),
  subShiftDown: em(subscriptShiftDown),
  supSubGapMin: em(subSuperscriptGapMin),
  spaceAfterScript: em(spaceAfterScript),
  upperLimitGapMin: em(upperLimitGapMin),
  upperLimitBaselineRiseMin: em(upperLimitBaselineRiseMin),
  lowerLimitGapMin: em(lowerLimitGapMin),
  lowerLimitBaselineDropMin: em(lowerLimitBaselineDropMin),
  stackGapMin: em(stackGapMin),
  radicalVerticalGap: em(radicalVerticalGap),
  radicalRuleThickness: em(radicalRuleThickness),
  radicalExtraAscender: em(radicalExtraAscender),
  radicalKernBeforeDegree: em(radicalKernBeforeDegree),
  radicalKernAfterDegree: em(radicalKernAfterDegree),
  radicalDegreeRaise: radicalDegreeBottomRaisePercent / 100,
  overbarVerticalGap: em(overbarVerticalGap),
  overbarRuleThickness: em(overbarRuleThickness),
  overbarExtraAscender: em(overbarExtraAscender),
  underbarVerticalGap: em(underbarVerticalGap),
  accentBaseHeight: em(accentBaseHeight),
  displayOperatorMinHeight: em(displayOperatorMinHeight),
  delimitedSubFormulaMinHeight: em(delimitedSubFormulaMinHeight),
  mathLeading: em(mathLeading),
};

const body = `// GENERATED by scripts/extract-math-constants.mjs from STIX Two Math's OpenType
// MATH table — do not edit by hand. Values are fractions of the em (font size),
// multiplied by the font size at layout time (see mathFont.ts).

export interface StixMathConstants {
${Object.keys(data).map((k) => `  ${k}: number;`).join("\n")}
}

export const STIX_MATH_CONSTANTS: StixMathConstants = ${JSON.stringify(data, null, 2)};
`;

writeFileSync(OUT, body);
console.log(`unitsPerEm=${unitsPerEm}`);
console.log(`wrote ${OUT}`);
console.log(JSON.stringify(data, null, 2));
