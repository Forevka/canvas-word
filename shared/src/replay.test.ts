import { describe, expect, it } from "vitest";
import type { Change } from "./change";
import type { Document, Paragraph, SectionProps } from "./model/document";
import { textOfRuns } from "./model/text";
import type { Op } from "./model/ops";
import { serializeDocument } from "./persist/serialize";
import { reconstruct } from "./replay";

const CHAR = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA = { align: "left" as const, lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };

const section = (): SectionProps => ({ pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } });
const para = (id: string, text: string): Paragraph => ({ kind: "paragraph", id, revision: 0, runs: [{ text, style: CHAR }], style: PARA });

const insert = (blockId: string, offset: number, text: string, seq: number): Change => ({
  id: `ch-${seq}`,
  docId: "d1",
  baseVersion: seq,
  seq,
  siteId: "siteA",
  origin: "typing",
  ts: 0,
  ops: [{ type: "insertText", at: { blockId, offset }, text } as Op],
});

const bodyText = (doc: Document): string => textOfRuns((doc.blocks[0] as Paragraph).runs);

describe("reconstruct (base snapshot + ordered change log)", () => {
  it("folds the change log onto the base to reach the latest version", () => {
    const base = serializeDocument({ section: section(), blocks: [para("p1", "hello")] });
    const changes = [insert("p1", 5, " world", 0), insert("p1", 0, ">> ", 1)];
    expect(bodyText(reconstruct(base, changes))).toBe(">> hello world");
  });

  it("orders by seq regardless of input order", () => {
    const base = serializeDocument({ section: section(), blocks: [para("p1", "hello")] });
    const changes = [insert("p1", 0, ">> ", 1), insert("p1", 5, " world", 0)]; // reversed
    expect(bodyText(reconstruct(base, changes))).toBe(">> hello world");
  });

  it("time-travels with uptoSeq (inclusive)", () => {
    const base = serializeDocument({ section: section(), blocks: [para("p1", "hello")] });
    const changes = [insert("p1", 5, " world", 0), insert("p1", 0, ">> ", 1)];
    expect(bodyText(reconstruct(base, changes, 0))).toBe("hello world"); // only seq 0
    expect(bodyText(reconstruct(base, changes, -1))).toBe("hello"); // nothing applied
  });

  it("does not mutate the base snapshot (replayable repeatedly)", () => {
    const base = serializeDocument({ section: section(), blocks: [para("p1", "hello")] });
    const changes = [insert("p1", 5, "!", 0)];
    reconstruct(base, changes);
    reconstruct(base, changes);
    expect(bodyText(reconstruct(base, changes))).toBe("hello!"); // stable across replays
    expect((base.doc.blocks[0] as Paragraph).runs[0]!.text).toBe("hello"); // base untouched
  });

  it("replays a multi-op change (split) faithfully", () => {
    const base = serializeDocument({ section: section(), blocks: [para("p1", "abcdef")] });
    const splitChange: Change = {
      id: "ch-split",
      docId: "d1",
      baseVersion: 0,
      seq: 0,
      siteId: "siteA",
      origin: "command",
      ts: 0,
      ops: [{ type: "splitParagraph", at: { blockId: "p1", offset: 3 }, newBlockId: "siteA-1" } as Op],
    };
    const doc = reconstruct(base, [splitChange]);
    expect(doc.blocks).toHaveLength(2);
    expect(textOfRuns((doc.blocks[0] as Paragraph).runs)).toBe("abc");
    expect(textOfRuns((doc.blocks[1] as Paragraph).runs)).toBe("def");
  });
});
