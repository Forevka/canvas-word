// Alignment keys for block-level LCS. Two blocks with the SAME key are treated
// as "the same block" by the aligner and get an intra-block diff; different keys
// show up as add/remove. Keys are intentionally coarse (content shape, not
// formatting) so a paragraph whose text is unchanged but whose STYLE changed
// still aligns as the same block — surfacing a format change rather than a
// delete+insert. Pure + DOM-free.

import type { Block, TableBlock } from "../model/document";
import { textOfRuns } from "../model/text";

/** Coarse "same content" key for block-level alignment. */
export function blockKey(b: Block): string {
  switch (b.kind) {
    case "paragraph":
      return `p:${textOfRuns(b.runs)}`;
    case "image":
      // Content-addressed bytes when available; else the (session-local) src.
      return `img:${b.mediaId ?? b.src}`;
    case "equation":
      return `eq:${JSON.stringify(b.equation)}`;
    case "table":
      return `tbl:${tableSignature(b)}`;
  }
}

/** Row×col grid shape plus a coarse concatenation of cell text — enough to align
 *  an unchanged table across the two documents without keying on cell ids. */
function tableSignature(t: TableBlock): string {
  const rows = t.rows.length;
  const cols = t.rows[0]?.cells.length ?? 0;
  const text = t.rows
    .map((r) => r.cells.map((c) => c.blocks.map((b) => (b.kind === "paragraph" ? textOfRuns(b.runs) : b.kind)).join("|")).join("¦"))
    .join("§");
  return `${rows}x${cols}:${text}`;
}
