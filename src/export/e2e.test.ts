// End-to-end against the real 8.9MB appraisal report (when present at the repo
// root). Import -> export PDF + DOCX in Node -> validate. Skipped if the fixture
// isn't available (it isn't committed), so CI without it still passes.

import { existsSync, readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import type { Block, Document } from "../model/document";
import { runImport } from "../import/docx/pipeline";
import { runExport } from "./pipeline";

const REPORT = fileURLToPath(
  new URL("../../RenderedReportsDocx_Report-AppraiseRequest-42103-Version-101.docx", import.meta.url),
);
const present = existsSync(REPORT);
const maybe = present ? describe : describe.skip;

async function resolveImages(doc: Document): Promise<Record<string, Uint8Array>> {
  const srcs = new Set<string>();
  const walk = (blocks: Block[]): void => {
    for (const b of blocks) {
      if (b.kind === "image") srcs.add(b.src);
      else if (b.kind === "table") for (const r of b.rows) for (const c of r.cells) walk(c.blocks);
    }
  };
  walk(doc.blocks);
  const out: Record<string, Uint8Array> = {};
  await Promise.all(
    [...srcs].map(async (s) => {
      try {
        out[s] = new Uint8Array(await (await fetch(s)).arrayBuffer());
      } catch {
        /* leave unresolved */
      }
    }),
  );
  return out;
}

maybe("E2E: real 8.9MB report", () => {
  it("imports, then exports to PDF and DOCX with faithful structure", async () => {
    const { doc } = runImport(new Uint8Array(readFileSync(REPORT)));
    expect(doc.blocks.length).toBeGreaterThan(500);
    const images = await resolveImages(doc);

    const pdf = await runExport(doc, "pdf", images);
    expect(pdf.bytes[0]).toBe(0x25); // %PDF
    const pages = (Buffer.from(pdf.bytes).toString("latin1").match(/\/Type\s*\/Page(?![s])/g) ?? []).length;
    expect(pages).toBeGreaterThan(50);

    const docx = await runExport(doc, "docx", images);
    expect(docx.bytes[0]).toBe(0x50); // PK
    // Round-trip: the exported docx re-imports to the same block count.
    const reimported = runImport(docx.bytes).doc;
    expect(reimported.blocks.length).toBe(doc.blocks.length);
  }, 30_000);
});
