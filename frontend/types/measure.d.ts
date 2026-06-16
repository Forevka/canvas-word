// Public type for @forevka/wordcanvas/export/measure — installs a DOM-free,
// fontkit-backed text-measurement context over the bundled clone fonts, so the
// layout engine (and PDF rendering) run headless in the export worker and on Node.
// MUST be awaited before exporting / laying out / recalculating a TOC.

/** Install the headless measurement host. Idempotent; safe to await repeatedly. */
export declare function installMeasureHost(): Promise<void>;
