// Public block-type registry — add a NEW document block type via one registration
// with a { measure, paint, toOOXML? } contract, instead of hand-editing the ~8
// dispatch sites that enumerate the built-in Block union.
//
// Scope (v1): canvas-drawn, ATOMIC (measures to a single box, places whole like
// an image/equation), NON-EDITABLE (no caret — skipped by geometry indexing,
// selected/deleted as a whole block), and JSON-serializable (snapshot serialize
// is free). A custom block in the model is `{ kind: "custom", customType, data,
// id, revision }` (see shared CustomBlock); `customType` selects the renderer
// registered here. The registry is a module singleton, shared by the layout
// engine (measure), the painter (paint), and the .docx exporter (toOOXML).
//
// NOT in v1: editable custom blocks (internal caret), table-cell nesting
// (body-level only), resize handles, and OOXML import round-trip.

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
