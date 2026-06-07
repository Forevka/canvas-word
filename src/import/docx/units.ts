// OOXML unit conversions → CSS px at 96dpi. All pure.
//
//   twips        1/20 pt          margins, indents, spacing
//   half-points  w:sz font size
//   EMU          914400/inch      image extents
//   w:line       240ths of a line when lineRule="auto"

/** twips → px: tw/20 pt × 96/72 = tw/15. */
export const twipsToPx = (tw: number): number => tw / 15;

/** half-points → px: hp/2 pt × 96/72 = hp × 2/3. */
export const halfPointsToPx = (hp: number): number => (hp * 2) / 3;

/** EMU → px: 914400 EMU/inch ÷ 96 px/inch = 9525 EMU/px. */
export const emuToPx = (emu: number): number => emu / 9525;

/** w:spacing w:line with lineRule="auto" (or absent): 240ths of a single line. */
export const lineAutoToMultiplier = (line: number): number => line / 240;

/** Round to 2 decimals — keeps layout math in floats but model JSON readable. */
export const round2 = (v: number): number => Math.round(v * 100) / 100;
