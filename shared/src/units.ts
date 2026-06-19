// Measurement units. The document model measures everything in CSS pixels at
// 96 dpi (Word's default screen resolution), so layout/paint can use px
// directly. These constants + helpers convert to/from the points and inches
// that docx and the UI surface — keep the magic 96/72 out of call sites.

/** CSS pixels per inch at the model's fixed 96 dpi. */
export const PX_PER_INCH = 96;
/** Typographic points per inch. */
export const PT_PER_INCH = 72;

/** px → pt (×0.75 at 96 dpi). */
export const pxToPt = (px: number): number => (px * PT_PER_INCH) / PX_PER_INCH;
/** pt → px (×4/3 at 96 dpi). */
export const ptToPx = (pt: number): number => (pt * PX_PER_INCH) / PT_PER_INCH;
