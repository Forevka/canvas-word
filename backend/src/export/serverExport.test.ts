// Round-trips the server-side import and export pipelines in Node (no DB):
//   build a doc -> writeDocx -> parseAndStore (backend import) -> exportDoc.
// Proves the pure frontend pipelines run headless on the backend, that PDF
// layout works (installMeasureHost), and that embedded media survives the
// content-address bridge (bytes -> mediaId -> bytes).

import { runExport } from "@cw/frontend/export";
import { installMeasureHost } from "@cw/frontend/export/measure";
import {
  reconstruct,
  textOfRuns,
  type Block,
  type Document,
  type ImageBlock,
  type Paragraph,
  type SectionProps,
} from "@cw/shared";
import { beforeAll, describe, expect, it } from "vitest";
import { parseAndStore } from "../import/serverImport";
import { MemoryChangeStore } from "../store/memoryStore";
import { exportDoc } from "./serverExport";

const section = (): SectionProps => ({
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
});
const CHAR = {
  fontFamily: "Georgia",
  fontSizePx: 16,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#000",
};
const PARA = {
  align: "left" as const,
  lineHeight: 1.2,
  spaceBeforePx: 0,
  spaceAfterPx: 0,
  indentFirstLinePx: 0,
  indentLeftPx: 0,
};
const para = (id: string, text: string): Paragraph => ({
  kind: "paragraph",
  id,
  revision: 0,
  runs: [{ text, style: CHAR }],
  style: PARA,
});
const image = (id: string, src: string): ImageBlock => ({
  kind: "image",
  id,
  revision: 0,
  src,
  widthPx: 120,
  heightPx: 60,
  align: "center",
});

/** Flattened text across every paragraph (body + table cells). */
function allText(doc: Document): string {
  const out: string[] = [];
  const walk = (blocks: Block[] | undefined): void => {
    for (const b of blocks ?? []) {
      if (b.kind === "paragraph") out.push(textOfRuns(b.runs));
      else if (b.kind === "table") for (const row of b.rows) for (const cell of row.cells) walk(cell.blocks);
    }
  };
  walk(doc.blocks);
  return out.join("\n");
}

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4, 5]);

describe("server import/export round-trip", () => {
  beforeAll(async () => {
    await installMeasureHost();
  });

  it("parses an uploaded docx, stores it, and exports docx + pdf", async () => {
    const store = new MemoryChangeStore();

    // Build a .docx with text + an embedded image using the export pipeline.
    const source: Document = {
      section: section(),
      blocks: [para("p1", "Hello round trip"), image("i1", "img1")],
    };
    const { bytes: docxBytes } = await runExport(source, "docx", { img1: PNG });
    expect(docxBytes[0]).toBe(0x50); // "PK" — a real zip/docx
    expect(docxBytes[1]).toBe(0x4b);

    // Backend import: parse, content-address media, persist.
    const uploaded = await parseAndStore(docxBytes, "sample.docx", store, {
      id: "u1",
      firstName: "Up",
      lastName: "Loader",
    });
    expect(uploaded.docId).toBeTruthy();

    // The title is the filename without extension; creator attributed.
    expect(await store.getDocumentTitle(uploaded.docId)).toBe("sample");
    const docs = await store.listDocuments();
    expect(docs[0]?.createdBy).toBe("u1");

    // Reconstruct the stored doc: text preserved, image carries a mediaId, and
    // the bytes are retrievable from the media store.
    const snap = await store.getSnapshot(uploaded.docId);
    const stored = reconstruct(snap!.snapshot, await store.getChanges(uploaded.docId, 0));
    expect(allText(stored)).toContain("Hello round trip");

    const img = stored.blocks.find((b): b is ImageBlock => b.kind === "image");
    expect(img?.mediaId).toBeTruthy();
    const media = await store.getMedia(img!.mediaId!);
    expect(media).not.toBeNull();
    expect([...media!.bytes]).toEqual([...PNG]);

    // Server-side export to both formats.
    const docx = await exportDoc(uploaded.docId, "docx", store);
    expect(docx).not.toBeNull();
    expect(docx!.title).toBe("sample");
    expect(docx!.bytes[0]).toBe(0x50); // PK
    expect(docx!.bytes[1]).toBe(0x4b);

    const pdf = await exportDoc(uploaded.docId, "pdf", store);
    expect(pdf).not.toBeNull();
    // "%PDF" magic.
    expect(String.fromCharCode(...pdf!.bytes.slice(0, 4))).toBe("%PDF");
  });

  it("returns null when exporting a missing document", async () => {
    const store = new MemoryChangeStore();
    expect(await exportDoc("does-not-exist", "docx", store)).toBeNull();
  });
});
