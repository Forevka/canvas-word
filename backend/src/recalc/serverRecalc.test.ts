// Server-side TOC recalculation, headless in Node (no DB):
//   - stateless: docx bytes -> recalc -> docx bytes (recalcDocxBytes)
//   - persisted: stored doc -> recalc -> versioned Change (recalcAndPersist)
// A TOC entry "label\t<stale>" linked to a heading's bookmark gets its number
// rewritten to the heading's real layout page (1 here).

import { runExport } from "@forevka/wordcanvas/export";
import { installMeasureHost } from "@forevka/wordcanvas/export/measure";
import {
  reconstruct,
  serializeDocument,
  textOfRuns,
  type CharStyle,
  type Change,
  type Document,
  type Paragraph,
  type ParaStyle,
  type SectionProps,
} from "@cw/shared";
import { beforeAll, describe, expect, it } from "vitest";
import { MemoryChangeStore } from "../store/memoryStore";
import { recalcAndPersist, recalcDocxBytes } from "./serverRecalc";

const SECTION: SectionProps = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
};
const CHAR: CharStyle = {
  fontFamily: "Georgia",
  fontSizePx: 16,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#000",
};
const PARA: ParaStyle = {
  align: "left",
  lineHeight: 1.2,
  spaceBeforePx: 0,
  spaceAfterPx: 0,
  indentFirstLinePx: 0,
  indentLeftPx: 0,
};

/** A doc with one TOC entry ("Chapter One\t<stale>", linked) + its heading. */
function sampleDoc(staleNum: string): Document {
  const label = "Chapter One";
  const tocEntry: Paragraph = {
    kind: "paragraph",
    id: "toc1",
    revision: 0,
    runs: [
      { text: label, style: { ...CHAR, link: "#_Toc1" } },
      { text: `\t${staleNum}`, style: CHAR },
    ],
    style: PARA,
  };
  const heading: Paragraph = {
    kind: "paragraph",
    id: "h1",
    revision: 0,
    runs: [{ text: label, style: { ...CHAR, bold: true } }],
    style: { ...PARA, namedStyle: "heading1" },
  };
  return {
    section: SECTION,
    blocks: [tocEntry, heading],
    bookmarks: { _Toc1: { start: { blockId: "h1", offset: 0 }, end: { blockId: "h1", offset: label.length } } },
  };
}

describe("serverRecalc", () => {
  beforeAll(async () => {
    await installMeasureHost();
  });

  it("stateless: recalcDocxBytes rewrites a stale TOC number and returns a valid docx", async () => {
    const { bytes: docx } = await runExport(sampleDoc("9"), "docx");
    const out = await recalcDocxBytes(docx);
    expect(out.changed).toBe(1);
    expect(out.bytes[0]).toBe(0x50); // "PK" — a real zip/docx
    expect(out.bytes[1]).toBe(0x4b);
  });

  it("persisted: recalcAndPersist appends a Change and updates the stored TOC text", async () => {
    const store = new MemoryChangeStore();
    const created = await store.createDocument(serializeDocument(sampleDoc("9")));
    const before = await store.getChanges(created.docId, 0);

    const published: Change[] = [];
    const result = await recalcAndPersist(created.docId, store, (_id, change) => published.push(change));
    expect(result).not.toBeNull();
    expect(result!.changed).toBe(1);

    // A new versioned change was appended and broadcast.
    const after = await store.getChanges(created.docId, 0);
    expect(after.length).toBe(before.length + 1);
    expect(published).toHaveLength(1);

    // The reconstructed head now shows the heading's real page (1).
    const snap = await store.getSnapshot(created.docId);
    const head = reconstruct(snap!.snapshot, after);
    const entry = head.blocks.find((b) => b.id === "toc1") as Paragraph;
    expect(textOfRuns(entry.runs)).toBe("Chapter One\t1");
  });

  it("persisted: no change when the TOC is already current", async () => {
    const store = new MemoryChangeStore();
    const created = await store.createDocument(serializeDocument(sampleDoc("1")));
    const published: Change[] = [];
    const result = await recalcAndPersist(created.docId, store, (_id, change) => published.push(change));
    expect(result!.changed).toBe(0);
    expect(published).toHaveLength(0);
    expect(await store.getChanges(created.docId, 0)).toHaveLength(0);
  });

  it("persisted: returns null for a missing document", async () => {
    const store = new MemoryChangeStore();
    expect(await recalcAndPersist("nope", store, () => {})).toBeNull();
  });
});
