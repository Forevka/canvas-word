// Unit helpers for the builder API. The document model is px @96dpi everywhere
// (matching the import pipeline's twips→px conversion), so authors thinking in
// physical units convert at the API boundary and the model stays uniform.

/** Inches → px @96dpi. */
export const inches = (n: number): number => n * 96;

/** Centimeters → px @96dpi. */
export const cm = (n: number): number => (n * 96) / 2.54;

/** Points (1/72in) → px @96dpi. */
export const pt = (n: number): number => (n * 96) / 72;

/** Twips (1/20pt) → px @96dpi — same factor as the import pipeline. */
export const twips = (n: number): number => n / 15;

export interface PageSize {
  pageWidthPx: number;
  pageHeightPx: number;
}

/** Common page sizes in px @96dpi (Letter matches the editor's default doc). */
export const PAGE_SIZES = {
  Letter: { pageWidthPx: 816, pageHeightPx: 1056 },
  Legal: { pageWidthPx: 816, pageHeightPx: 1344 },
  A4: { pageWidthPx: 794, pageHeightPx: 1123 },
  A3: { pageWidthPx: 1123, pageHeightPx: 1587 },
  Tabloid: { pageWidthPx: 1056, pageHeightPx: 1632 },
} as const satisfies Record<string, PageSize>;

export type PageSizeName = keyof typeof PAGE_SIZES;
