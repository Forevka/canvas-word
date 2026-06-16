// Builder-only table-style presets. These are NOT a model concept — they are
// resolved into concrete cell formatting (shading, borders, run/para styles) at
// build time, so nothing style-id-like ever lands in the Document. A preset is the
// BASE; an explicit CellSpec value always wins (applied in TableBuilder.toBlock).
//
// Lives in its own module (no BuilderContext import) so both blockFactory.ts (which
// seeds the per-builder registry) and tableBuilder.ts (which applies presets) can
// import it without a runtime cycle.

import type { CellBorders, CharStyle } from "@cw/shared";

export interface TableStylePreset {
  /** Bold the first row and apply headerChar/headerShading to it. */
  headerRow?: boolean;
  /** Char formatting for the header row's runs (overrides the default bold). */
  headerChar?: Partial<CharStyle>;
  /** Header-row background fill (CSS color). */
  headerShading?: string;
  /** Per-edge borders applied to every cell that doesn't set its own. */
  borders?: CellBorders;
  /** Background fill for every cell that doesn't set its own. */
  shading?: string;
  /** Zebra striping: fill for alternating rows that don't set their own shading. */
  stripeShading?: string;
}

const GRID_EDGE = { color: "#d0d4d9", widthPx: 1 } as const;
const GRID_BORDERS: CellBorders = { top: GRID_EDGE, right: GRID_EDGE, bottom: GRID_EDGE, left: GRID_EDGE };

/** The presets every builder starts with (override or add via .tableStylePreset). */
export function builtinTableStyles(): Map<string, TableStylePreset> {
  return new Map<string, TableStylePreset>([
    ["plain", {}],
    ["grid", { borders: GRID_BORDERS }],
    ["headerBand", { headerRow: true, headerShading: "#1a73e8", headerChar: { color: "#ffffff", bold: true }, borders: GRID_BORDERS }],
    ["striped", { headerRow: true, headerShading: "#e8f0fe", headerChar: { bold: true }, stripeShading: "#f6f9fe" }],
  ]);
}
