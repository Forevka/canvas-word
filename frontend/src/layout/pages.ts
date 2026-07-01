// Page-level query: "what's on page N". Pages do NOT exist in the document model
// — they are produced by the layout engine, so this is the one part of the
// query/edit API that lives in the frontend (layout) package rather than the
// shared model core. It runs a layout pass and projects the LayoutTree into a
// small, serializable page map (safe to marshal across the C# boundary).
//
// Preconditions mirror rendering/export: text measurement (pretext) must be
// initialized in the host (it is in the browser editor and in the headless
// export/measure pipelines).

import type { Document } from "@cw/shared";
import { createLayoutEngine } from "./engine";

export interface PageInfo {
  /** 0-based sequence order. */
  index: number;
  /** Displayed page number (honors section pageNumberStart). */
  number: number;
  widthPx: number;
  heightPx: number;
  marginPx: { top: number; right: number; bottom: number; left: number };
  /** Body content-box top/bottom (differ from margins when a tall band pushes the body). */
  contentTopPx: number;
  contentBottomPx: number;
  /** Page fill ("#rrggbb"); absent = white. */
  pageColorHex?: string;
  /** Top-level body block ids placed on this page, in reading order. A block split
   *  across a page boundary appears on each page it covers. */
  blockIds: string[];
}

export interface GetPagesOptions {
  /** Custom-font registry to make active for the layout pass (see createLayoutEngine). */
  fontRegistry?: Parameters<typeof createLayoutEngine>[0];
  /** Per-engine CJK/script fallback + locale (see createLayoutEngine). */
  engineOptions?: Parameters<typeof createLayoutEngine>[1];
}

/** Lay the document out and return a serializable per-page map — the answer to
 *  "what's on page 8–9". Page numbers can shift after edits, so re-run this after
 *  mutating the document. */
export function getPages(doc: Document, options: GetPagesOptions = {}): PageInfo[] {
  const engine = createLayoutEngine(options.fontRegistry, options.engineOptions);
  const tree = engine.layout(doc);
  return tree.pages.map((page) => {
    const seen = new Set<string>();
    const blockIds: string[] = [];
    for (const block of page.blocks) {
      if (!seen.has(block.blockId)) {
        seen.add(block.blockId);
        blockIds.push(block.blockId);
      }
    }
    const info: PageInfo = {
      index: page.index,
      number: page.number,
      widthPx: page.widthPx,
      heightPx: page.heightPx,
      marginPx: page.marginPx,
      contentTopPx: page.contentTopPx,
      contentBottomPx: page.contentBottomPx,
      blockIds,
    };
    if (page.pageColorHex !== undefined) info.pageColorHex = page.pageColorHex;
    return info;
  });
}

/** The displayed page number a top-level block first appears on, or null if it
 *  isn't placed. For many lookups, call `getPages` once and scan the result
 *  instead of calling this repeatedly (each call runs a full layout pass). */
export function pageOfBlock(doc: Document, blockId: string, options?: GetPagesOptions): number | null {
  for (const page of getPages(doc, options)) {
    if (page.blockIds.includes(blockId)) return page.number;
  }
  return null;
}
