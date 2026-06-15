// px (CSS, 96dpi) -> OOXML units. Exact inverses of src/import/docx/units.ts.

// A non-finite value (NaN/Infinity from an undefined style field) would serialize
// to an invalid w:val="NaN" and corrupt the .docx. Coerce to 0 so output stays
// well-formed — these are pure unit conversions, so NaN never legitimately occurs.
const finite = (n: number): number => (Number.isFinite(n) ? n : 0);

/** px -> twips: px × 15 (twip = 1/20pt, px = tw/15). */
export const pxToTwips = (px: number): number => Math.round(finite(px) * 15);

/** px -> half-points: px × 1.5 (w:sz; px = hp × 2/3). */
export const pxToHalfPoints = (px: number): number => Math.round(finite(px) * 1.5);

/** px -> EMU: px × 9525 (image extents; 914400 EMU/in ÷ 96). */
export const pxToEmu = (px: number): number => Math.round(finite(px) * 9525);

/** px -> eighth-points: px × 6 (border w:sz; px = sz/6). */
export const pxToEighthPoints = (px: number): number => Math.max(2, Math.round(finite(px) * 6));

/** line-height multiplier -> w:line 240ths (lineRule="auto"). */
export const multiplierToLine = (m: number): number => Math.round(finite(m) * 240);
