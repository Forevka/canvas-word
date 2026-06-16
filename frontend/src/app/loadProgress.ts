// First-load progress contract, shared by the two emit sites: the WordCanvas
// constructor (which brackets the lazy editorApp chunk download) and the editor
// mount (which streams the bundled fonts). Kept in one place so the phase weights
// stay consistent and a single `onLoadProgress` callback sees a monotonic 0..1.

/** First-load progress, surfaced to embedders via `WordCanvasOptions.onLoadProgress`. */
export interface LoadProgress {
  /** Coarse stage: "bundle" = downloading the editor JS chunk (indeterminate —
   *  a dynamic import() can't be byte-measured); "fonts" = fetching the bundled
   *  fonts (measurable); "ready" = mount complete. */
  phase: "bundle" | "fonts" | "ready";
  /** Overall startup progress, 0..1, monotonic. Drive a progress bar with this. */
  percent: number;
  /** Bytes fetched / total for the current measurable phase (fonts). Both 0 when
   *  the phase is indeterminate (bundle). Informational. */
  loaded: number;
  total: number;
}

/** Share of the bar reserved for the (indeterminate) JS-chunk download. The bundle
 *  phase reports 0 while downloading and snaps to this when the chunk resolves;
 *  the fonts phase then fills the remaining `1 - BUNDLE_SHARE`. */
export const BUNDLE_SHARE = 0.15;

/** Build the fonts-phase progress object, mapping streamed font bytes onto the
 *  `BUNDLE_SHARE → 1.0` slice of the overall bar. */
export function fontsProgress(loaded: number, total: number): LoadProgress {
  const frac = total > 0 ? loaded / total : 1;
  return { phase: "fonts", loaded, total, percent: BUNDLE_SHARE + (1 - BUNDLE_SHARE) * frac };
}
