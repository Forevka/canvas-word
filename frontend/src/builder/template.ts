// fromTemplate glue: a template .docx contributes its STYLES — stylesheet,
// list definitions, page setup, header/footer bands — while body content is
// discarded by default (the builder composes fresh content against those
// styles). Runs the import pipeline in media-collector mode (no blob: URLs,
// works in Node) and inlines collected image bytes as data: URLs so the
// resulting Document is portable across editor, browser worker, and Node.

import type { Block, Document } from "@cw/shared";
import { BAND_CONTAINERS } from "@cw/shared";
import { runImport } from "../import/docx/pipeline";
import type { ImportWarning } from "../import/docx/types";
import { bytesToDataUrl } from "./media";

export function prepareTemplate(
  docx: ArrayBuffer | Uint8Array,
  keepBody: boolean,
): { doc: Document; warnings: ImportWarning[] } {
  const bytes = docx instanceof Uint8Array ? docx : new Uint8Array(docx);
  const result = runImport(bytes, undefined, { collectMediaBytes: true });
  const doc = result.doc;

  // cw-media:N → data: URL (collector mode never minted blob: URLs).
  const remap = new Map(result.media.map((m) => [m.src, bytesToDataUrl(m.bytes, m.mime)]));
  if (remap.size > 0) {
    const rewrite = (blocks: Block[] | undefined): void => {
      for (const b of blocks ?? []) {
        if (b.kind === "image") {
          const url = remap.get(b.src);
          if (url) b.src = url;
        } else if (b.kind === "table") {
          for (const row of b.rows) for (const cell of row.cells) rewrite(cell.blocks);
        }
      }
    };
    rewrite(doc.blocks);
    for (const band of BAND_CONTAINERS) rewrite(doc.section[band]);
  }

  if (!keepBody) {
    doc.blocks = [];
    // Footnotes are referenced from body runs — with the body gone they are
    // unreachable. Bookmarks and content controls may live in bands, so keep
    // only the ones whose anchors survived.
    delete doc.footnotes;
    pruneBodyAnchored(doc);
  }

  return { doc, warnings: result.warnings };
}

/** Drop bookmarks/sdts whose anchor blocks/runs no longer exist. */
function pruneBodyAnchored(doc: Document): void {
  const blockIds = new Set<string>();
  const sdtIds = new Set<string>();
  const visit = (blocks: Block[] | undefined): void => {
    for (const b of blocks ?? []) {
      blockIds.add(b.id);
      for (const id of b.sdtPath ?? []) sdtIds.add(id);
      if (b.kind === "paragraph") {
        for (const r of b.runs) for (const id of r.style.sdtPath ?? []) sdtIds.add(id);
      } else if (b.kind === "table") {
        for (const row of b.rows) for (const cell of row.cells) visit(cell.blocks);
      }
    }
  };
  visit(doc.blocks);
  for (const band of BAND_CONTAINERS) visit(doc.section[band]);

  if (doc.bookmarks) {
    for (const [name, range] of Object.entries(doc.bookmarks)) {
      if (!blockIds.has(range.start.blockId) || !blockIds.has(range.end.blockId)) delete doc.bookmarks[name];
    }
    if (Object.keys(doc.bookmarks).length === 0) delete doc.bookmarks;
  }
  if (doc.sdts) {
    for (const id of Object.keys(doc.sdts)) {
      if (!sdtIds.has(id)) delete doc.sdts[id];
    }
    if (Object.keys(doc.sdts).length === 0) delete doc.sdts;
  }
}
