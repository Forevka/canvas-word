// Length unit conversion for the Page Layout dialog. The model is native CSS px
// at 96 dpi everywhere; these helpers convert ONLY at the input boundary so the
// user can type inches or centimetres. 1 in = 96 px; 1 cm = 96/2.54 px.

export type LengthUnit = "in" | "cm";

const CM_PER_IN = 2.54;

export const inToPx = (inch: number): number => inch * 96;
export const pxToIn = (px: number): number => px / 96;
export const cmToPx = (cm: number): number => (cm / CM_PER_IN) * 96;
export const pxToCm = (px: number): number => (px / 96) * CM_PER_IN;

export const pxToUnit = (px: number, u: LengthUnit): number => (u === "in" ? pxToIn(px) : pxToCm(px));
export const unitToPx = (v: number, u: LengthUnit): number => (u === "in" ? inToPx(v) : cmToPx(v));

/** Format a px length for display in the given unit, rounded to 2 decimals. */
export const formatUnit = (px: number, u: LengthUnit): string => {
  const v = pxToUnit(px, u);
  return (Math.round(v * 100) / 100).toString();
};
