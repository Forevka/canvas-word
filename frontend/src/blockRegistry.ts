// Public block-type registry — add a NEW document block type via one registration
// with a { measure, paint, toOOXML? } contract, instead of hand-editing the ~8
// dispatch sites that enumerate the built-in Block union.
//
// The block contract: a custom block is a canvas-drawn ATOMIC object — it measures
// to a single box and places whole (like an image/equation), and holds only a
// JSON-serializable `data` payload (so snapshot serialize/paste is free). In the
// model it is `{ kind: "custom", customType, data, id, revision }` (see shared
// CustomBlock); `customType` selects the renderer registered here. It lays out and
// paints anywhere blocks go (body, table cells, header/footer bands), and is a
// first-class OBJECT in the editor: click to select (a plain, non-resizable frame,
// like an equation), Delete / right-click ▸ Delete to remove, undo/redo like any
// edit. The registry is a module singleton, shared by the layout engine (measure),
// the on-screen painter, the PDF exporter, and the .docx exporter.
//
// PDF export renders the SAME `paint(ctx, box, data)` — it's replayed through a
// Canvas2D→pdfkit vector shim (see export/pdf/canvasToPdf.ts), so a block renders
// crisply in PDF with no extra wiring and no browser: it works headlessly on the
// Node backend and in the bare-V8 ClearScript host. The one caveat is the browser
// EDITOR's own PDF export, which normally runs in a worker with no registry — for
// a document containing a registered custom block it runs the export inline on the
// main thread instead (see export/exportDocument.ts). A block whose paint uses ops
// the vector shim can't translate can set `pdf: "raster"` (a reserved seam).
//
// By design a custom block carries no internal caret — its content is drawn, not
// text-edited (that is what paragraphs and content controls are for), and its size
// is owned by `measure()` rather than interactive resize. It has no native OOXML,
// so `.docx` export is lossy unless the type supplies `toOOXML` (see below), and a
// custom block never originates from an imported `.docx`.

/** Measurement context passed to a custom block's `measure`. */
export interface CustomBlockMeasureCtx {
  /** The content width (document px) available to the block on its page/section. */
  width: number;
}

/** What `measure` returns — the block's atomic box height for the given width. */
export interface CustomBlockSize {
  height: number;
}

/** The box a custom block is painted into (document px). The canvas context is
 *  already translated to the block's top-left, so paint in local coordinates:
 *  `[0, width] × [0, height]`. */
export interface CustomBlockBox {
  width: number;
  height: number;
}

/** A registered custom block type. `type` is the key stored on a block's
 *  `customType`. */
export interface CustomBlockType {
  /** Unique key; matches `CustomBlock.customType`. */
  type: string;
  /** Measure the block: given the available content width, return its height.
   *  Called during layout; must be a pure function of `data` + width (it feeds
   *  the revision-keyed layout cache). */
  measure(data: unknown, ctx: CustomBlockMeasureCtx): CustomBlockSize;
  /** Paint the block. The canvas context is translated to the block's top-left
   *  (draw in `[0, box.width] × [0, box.height]`) and clipped to the box. Draw
   *  only — never measure text here (paint-never-measures). */
  paint(ctx: CanvasRenderingContext2D, box: CustomBlockBox, data: unknown): void;
  /** OPTIONAL: serialize the block to WordprocessingML for `.docx` export
   *  (return the XML for one block, e.g. a `<w:p>…</w:p>` or `<w:tbl>…`). Omit and
   *  export is lossy — a placeholder empty paragraph is emitted and a warning is
   *  reported. */
  toOOXML?(data: unknown): string;
  /** How to render this block into PDF. Default `"vector"`: `paint()` is replayed
   *  through a Canvas2D→pdfkit shim as crisp vector ops (works headlessly — Node /
   *  ClearScript — with no native canvas). `"raster"` is a reserved escape hatch
   *  for blocks whose `paint()` uses ops the vector shim can't translate (shadows,
   *  filters, pattern fills, image compositing); it will rasterize via a headless
   *  canvas. Until the raster path ships, a `"raster"` block reserves its space in
   *  PDF and reports a warning. */
  pdf?: "vector" | "raster";
}

const registry = new Map<string, CustomBlockType>();

/** Register a custom block type. Returns a disposer that unregisters it. A
 *  duplicate `type` overwrites the previous registration (with a console warning),
 *  so hot-reload / re-registration is safe. */
export function registerBlockType(def: CustomBlockType): () => void {
  if (registry.has(def.type)) {
    console.warn(`[canvas-word] registerBlockType: overwriting existing block type "${def.type}"`);
  }
  registry.set(def.type, def);
  return () => {
    if (registry.get(def.type) === def) registry.delete(def.type);
  };
}

/** Look up a registered block type, or undefined. */
export function getBlockType(type: string): CustomBlockType | undefined {
  return registry.get(type);
}

/** Whether a block type is registered. */
export function hasBlockType(type: string): boolean {
  return registry.has(type);
}

/** Test-only: clear all registrations. */
export function _clearBlockTypes(): void {
  registry.clear();
}
