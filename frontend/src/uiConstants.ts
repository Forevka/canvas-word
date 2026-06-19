// Editor UI tuning constants shared across more than one module (so a value
// isn't retyped per call site). Single-use UI literals stay local to their
// module; this file is only for the cross-file ones.

/** Multiplicative zoom step for the +/- controls and Ctrl+wheel. */
export const ZOOM_STEP = 1.1;
/** Absolute zoom clamp the renderer enforces (the slider stays well inside). */
export const ZOOM_MIN = 0.25;
export const ZOOM_MAX = 5;

/** Indent / outdent step in px (0.375in at 96 dpi), per toolbar press. */
export const INDENT_STEP_PX = 36;
