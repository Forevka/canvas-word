// Regression guard for the FULL export round-trip of the default showcase document
// (the editor's initial state when opened with no docId). It answers: does the
// default document survive `export DOCX → open → export DOCX` with no losses or
// silent drift?
//
//   doc0 = sampleDoc()             the in-memory default
//   doc1 = import(export(doc0))    opened once
//   doc2 = import(export(doc1))    opened, saved, opened again
//
// The meaningful "no losses on open→save→open" property is that doc1 is a FIXED
// POINT: doc2 must equal doc1 (a stable save never mutates the document). We also
// pin a structural fingerprint so a regression that drops/duplicates content is
// caught with a legible diff rather than a giant object mismatch.
//
// This caught two real defects (see docx/bookmarkRoundtrip.test.ts): an
// inline bookmark whose end drifted 65→89 chars across saves (writer snapped a
// mid-run marker to the run boundary) and a 1-char offset shift from a footnote
// reference (importer counted the reference as 0-width while the model paints its
// number). Keep this test broad — it's the tripwire for export round-trip fidelity.

import { beforeAll, describe, expect, it } from "vitest";
import type { Block, Document } from "@cw/shared";
import { runImport } from "../import/docx/pipeline";
import { runExport } from "./pipeline";
import { installMeasureHost } from "./shared/measureHost";
import { sampleDoc } from "../model/sampleDoc";

async function resolveImages(doc: Document): Promise<Record<string, Uint8Array>> {
  const srcs = new Set<string>();
  const walk = (blocks: Block[] | undefined): void => {
    if (!blocks) return;
    for (const b of blocks) {
      if (b.kind === "image") srcs.add(b.src);
      else if (b.kind === "table") for (const r of b.rows) for (const c of r.cells) walk(c.blocks);
      else if (b.kind === "paragraph" && b.style.sectionBreak) {
        const p = b.style.sectionBreak.props;
        walk(p.header); walk(p.footer); walk(p.headerFirst);
        walk(p.headerEven); walk(p.footerFirst); walk(p.footerEven);
      }
    }
  };
  walk(doc.blocks);
  const s = doc.section;
  walk(s.header); walk(s.footer); walk(s.headerFirst);
  walk(s.headerEven); walk(s.footerFirst); walk(s.footerEven);
  const out: Record<string, Uint8Array> = {};
  await Promise.all(
    [...srcs].map(async (src) => {
      try {
        out[src] = new Uint8Array(await (await fetch(src)).arrayBuffer());
      } catch {
        /* leave unresolved — writer emits a placeholder + warning */
      }
    }),
  );
  return out;
}

// Image `src` (fresh blob:/synthetic URL per import) and any mediaId are volatile
// by construction; blank them before the deep compare so only real content drift
// registers.
function normalize(v: unknown): unknown {
  return JSON.parse(
    JSON.stringify(v, (key, val) => {
      if ((key === "src" || key === "mediaId") && typeof val === "string") return `<${key}>`;
      return val;
    }),
  );
}

// A quick structural census: a legible fingerprint of the document's content so a
// dropped run/table/field surfaces as a one-number diff.
function fingerprint(doc: Document): Record<string, number> {
  let paras = 0, runs = 0, tables = 0, cells = 0, images = 0, chars = 0, sectionBreaks = 0, equations = 0, shapes = 0;
  const walk = (blocks: Block[] | undefined): void => {
    if (!blocks) return;
    for (const b of blocks) {
      if (b.kind === "paragraph") {
        paras++;
        if (b.style.sectionBreak) {
          sectionBreaks++;
          const p = b.style.sectionBreak.props;
          walk(p.header); walk(p.footer); walk(p.headerFirst);
          walk(p.headerEven); walk(p.footerFirst); walk(p.footerEven);
        }
        for (const r of b.runs) { runs++; chars += r.text.length; }
      } else if (b.kind === "table") {
        tables++;
        for (const r of b.rows) for (const c of r.cells) { cells++; walk(c.blocks); }
      } else if (b.kind === "image") images++;
      else if (b.kind === "equation") equations++;
      else if (b.kind === "shape") { shapes++; walk(b.text?.blocks); }
    }
  };
  walk(doc.blocks);
  return {
    paras, runs, tables, cells, images, chars, sectionBreaks, equations, shapes,
    footnotes: Object.keys(doc.footnotes ?? {}).length,
    endnotes: Object.keys(doc.endnotes ?? {}).length,
    fields: Object.keys(doc.fields ?? {}).length,
    sdts: Object.keys(doc.sdts ?? {}).length,
    bookmarks: Object.keys(doc.bookmarks ?? {}).length,
  };
}

// Collect up to `limit` differing JSON paths between two normalized values.
function diffPaths(a: unknown, b: unknown, path = "", out: string[] = [], limit = 80): string[] {
  if (out.length >= limit || a === b) return out;
  const ta = typeof a;
  if (ta !== typeof b || a === null || b === null || ta !== "object") {
    out.push(`${path || "(root)"}: ${JSON.stringify(a)} → ${JSON.stringify(b)}`);
    return out;
  }
  if (Array.isArray(a) || Array.isArray(b)) {
    const aa = (a as unknown[]) ?? [], bb = (b as unknown[]) ?? [];
    if (aa.length !== bb.length) out.push(`${path}.length: ${aa.length} → ${bb.length}`);
    for (let i = 0; i < Math.max(aa.length, bb.length) && out.length < limit; i++) diffPaths(aa[i], bb[i], `${path}[${i}]`, out, limit);
    return out;
  }
  const ao = a as Record<string, unknown>, bo = b as Record<string, unknown>;
  for (const k of new Set([...Object.keys(ao), ...Object.keys(bo)])) {
    if (out.length >= limit) break;
    diffPaths(ao[k], bo[k], path ? `${path}.${k}` : k, out, limit);
  }
  return out;
}

describe("default showcase doc: export → open → export fidelity", () => {
  beforeAll(async () => {
    await installMeasureHost();
  });

  it("is a stable fixed point across open → save → open (doc1 === doc2)", async () => {
    const doc0 = sampleDoc();
    const doc1 = runImport((await runExport(doc0, "docx", await resolveImages(doc0))).bytes).doc;
    const doc2 = runImport((await runExport(doc1, "docx", await resolveImages(doc1))).bytes).doc;

    const fp1 = fingerprint(doc1);
    const fp2 = fingerprint(doc2);
    expect(fp2).toEqual(fp1); // no content dropped/duplicated on the second save

    const paths = diffPaths(normalize(doc1), normalize(doc2));
    expect(paths).toEqual([]); // byte-for-byte stable model — no silent drift
  }, 60_000);

  it("first export preserves the document's structural census", async () => {
    // Pins the shape of the default doc after one round-trip. Benign, EXPECTED
    // normalizations vs the in-memory model: adjacent same-style runs coalesce
    // (fewer runs, identical char count), and the TOC's implicit PAGEREF anchor
    // bookmarks materialize. If these numbers change, confirm the cause is a
    // deliberate change to sampleDoc()/the pipeline before updating them.
    const doc0 = sampleDoc();
    const doc1 = runImport((await runExport(doc0, "docx", await resolveImages(doc0))).bytes).doc;
    expect(fingerprint(doc1)).toEqual({
      paras: 334,
      runs: 493,
      tables: 11,
      cells: 203,
      images: 4,
      chars: 24702,
      sectionBreaks: 2,
      equations: 5,
      shapes: 23,
      footnotes: 1,
      endnotes: 1,
      fields: 8,
      sdts: 13,
      bookmarks: 16,
    });
  }, 60_000);
});
