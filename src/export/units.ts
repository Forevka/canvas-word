// px (CSS, 96dpi) -> OOXML units. Exact inverses of src/import/docx/units.ts.

/** px -> twips: px × 15 (twip = 1/20pt, px = tw/15). */
export const pxToTwips = (px: number): number => Math.round(px * 15);

/** px -> half-points: px × 1.5 (w:sz; px = hp × 2/3). */
export const pxToHalfPoints = (px: number): number => Math.round(px * 1.5);

/** px -> EMU: px × 9525 (image extents; 914400 EMU/in ÷ 96). */
export const pxToEmu = (px: number): number => Math.round(px * 9525);

/** px -> eighth-points: px × 6 (border w:sz; px = sz/6). */
export const pxToEighthPoints = (px: number): number => Math.max(2, Math.round(px * 6));

/** line-height multiplier -> w:line 240ths (lineRule="auto"). */
export const multiplierToLine = (m: number): number => Math.round(m * 240);
