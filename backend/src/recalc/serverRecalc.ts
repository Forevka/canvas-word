// Server-side TOC recalculation: run THIS project's layout engine headless to give
// each TOC entry its heading's real page number — the thing a renderer-less caller
// (e.g. a C# pipeline that emits a .docx with a TOC field) cannot compute itself.
//
// installMeasureHost() injects the DOM-free fontkit measure context, so layout runs
// in Node. The recalc only edits TOC entries' trailing page-number text, so the
// layout never needs image pixels (ImageBlock carries its own width/height) — unlike
// export, no media hydration is required here.
//
// Two entry points: a stateless bytes-in/bytes-out path for throwaway recalcs, and a
// docId path that persists the rewrite as an OT Change so it's versioned and pushed
// live to connected editors.

import { runImport } from "@forevka/wordcanvas/import";
import { runExport } from "@forevka/wordcanvas/export";
import { installMeasureHost } from "@forevka/wordcanvas/export/measure";
import { computeTocEditsWithLayout, recalcToc, tocEditsToOps } from "@forevka/wordcanvas/recalc";
import { reconstruct, type Change } from "@cw/shared";
import { randomUUID } from "node:crypto";
import type { ChangeStore } from "../store/ChangeStore";

// Layout (and the stateless path's export) is CPU-heavy; cap concurrency so a burst
// can't starve the event loop / collab traffic. Mirrors serverExport.ts.
const MAX_CONCURRENT = Number(process.env.RECALC_CONCURRENCY ?? process.env.EXPORT_CONCURRENCY ?? 2);
let active = 0;
const queue: (() => void)[] = [];

function acquire(): Promise<void> {
  if (active < MAX_CONCURRENT) {
    active++;
    return Promise.resolve();
  }
  return new Promise<void>((resolve) => queue.push(resolve));
}

function release(): void {
  active--;
  const next = queue.shift();
  if (next) {
    active++;
    next();
  }
}

export interface RecalcDocxResult {
  bytes: Uint8Array;
  /** Number of TOC entries whose page number changed. */
  changed: number;
  warnings: { code: string; message: string }[];
}

/** Stateless: import a .docx, recalculate its TOC page numbers, export it back.
 *  No persistence — for a render pipeline that just wants a corrected file. */
export async function recalcDocxBytes(bytes: Uint8Array): Promise<RecalcDocxResult> {
  // collectMediaBytes stamps each image with a synthetic src ("cw-media:N") and
  // hands back the bytes; the exporter keys embedded images by that same src, so we
  // pass them straight through (no content-addressing round-trip needed here).
  const result = runImport(bytes, undefined, { collectMediaBytes: true });
  const images: Record<string, Uint8Array> = {};
  for (const m of result.media) images[m.src] = m.bytes;

  await acquire();
  try {
    await installMeasureHost();
    const { doc, changed } = recalcToc(result.doc);
    const { bytes: out } = await runExport(doc, "docx", images);
    return { bytes: out, changed, warnings: result.warnings };
  } finally {
    release();
  }
}

export interface RecalcPersistResult {
  changed: number;
  /** New head version after the recalc change (unchanged when nothing was stale). */
  version: number;
}

/** docId path: reconstruct the live head, recalculate, and persist the rewrite as an
 *  OT Change (versioned, audit-visible, broadcast to connected editors). Uses the
 *  live (un-baked) document and layout — matching what editors see and the editor's
 *  own "recalculate TOC" command; the export endpoint handles track-changes baking
 *  separately. Returns null if the document doesn't exist. */
export async function recalcAndPersist(
  docId: string,
  store: ChangeStore,
  publish: (docId: string, change: Change) => void,
): Promise<RecalcPersistResult | null> {
  const snap = await store.getSnapshot(docId);
  if (!snap) return null;
  const changes = await store.getChanges(docId, snap.version);
  const head = snap.version + changes.length;
  const doc = reconstruct(snap.snapshot, changes);

  await acquire();
  let ops;
  try {
    await installMeasureHost();
    ops = tocEditsToOps(computeTocEditsWithLayout(doc));
  } finally {
    release();
  }
  if (ops.length === 0) return { changed: 0, version: head };

  const change: Change = {
    id: randomUUID(),
    docId,
    baseVersion: head,
    siteId: "server-recalc",
    origin: "command",
    ts: Date.now(),
    ops,
    selectionAfter: null,
  };
  const accepted = await store.appendChange(docId, change);
  publish(docId, accepted);
  // Two ops (deleteRange + insertText) per changed entry.
  return { changed: ops.length / 2, version: accepted.seq ?? head + 1 };
}
