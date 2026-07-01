// Regression: an inline bookmark whose boundary falls inside a run must round-trip
// to the SAME covered text, save after save. Two defects used to corrupt it:
//
//  1. Writer snapped a mid-run bookmark marker to the run's END boundary (it only
//     flushed markers between runs). Once import coalesced the bookmarked run with
//     the next same-style run, the bookmark grew to swallow it — the span expanded
//     on every save until it hit the run boundary.
//  2. Importer counted a footnote reference as 0-width when resolving bookmark
//     offsets, but the model paints the reference's number (1 char) — so a bookmark
//     after a footnote drifted one character early.
//
// The showcase `sample` bookmark exercises BOTH: it sits after a footnote reference
// AND its end lands mid-run (its run and the following one share the body style, so
// they coalesce on import). It must keep covering exactly "bookmarked text".

import { beforeAll, describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import { runExport } from "../pipeline";
import { installMeasureHost } from "../shared/measureHost";
import { sampleDoc } from "../../model/sampleDoc";

// Text covered by a named bookmark: concat the home paragraph's runs and slice.
function coveredText(doc: Document, name: string): string {
  const bk = (doc.bookmarks ?? {})[name];
  if (!bk) return "<<missing>>";
  const p = doc.blocks.find((b): b is Paragraph => b.kind === "paragraph" && b.id === bk.start.blockId);
  if (!p) return "<<no-home>>";
  return p.runs.map((r) => r.text).join("").slice(bk.start.offset, bk.end.offset);
}

async function roundtrip(doc: Document): Promise<Document> {
  return runImport((await runExport(doc, "docx")).bytes).doc;
}

describe("inline bookmark span survives repeated open → save", () => {
  beforeAll(async () => {
    await installMeasureHost();
  });

  it("keeps the showcase 'sample' bookmark on exactly its text across saves", async () => {
    const doc0 = sampleDoc();
    expect(coveredText(doc0, "sample")).toBe("bookmarked text");

    const doc1 = await roundtrip(doc0);
    expect(coveredText(doc1, "sample")).toBe("bookmarked text"); // no footnote off-by-one

    const doc2 = await roundtrip(doc1);
    expect(coveredText(doc2, "sample")).toBe("bookmarked text"); // no mid-run boundary drift

    const doc3 = await roundtrip(doc2);
    expect(coveredText(doc3, "sample")).toBe("bookmarked text"); // stable fixed point
  }, 60_000);

  it("splits a run to place a mid-run bookmark boundary (no coalesce drift)", async () => {
    // A minimal reproduction independent of the showcase: two adjacent runs with
    // the SAME style, a bookmark covering only the first. On import the runs
    // coalesce; the next export must still land the bookmark end mid-(merged)-run.
    const CHAR: CharStyle = {
      fontFamily: "Times New Roman", fontSizePx: 16, bold: false, italic: false,
      underline: false, strikethrough: false, color: "#000",
    };
    const PARA: ParaStyle = {
      align: "left", lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 8,
      indentFirstLinePx: 0, indentLeftPx: 0,
    };
    const para: Paragraph = {
      kind: "paragraph", id: "p0", revision: 0, style: PARA,
      runs: [{ text: "keep this", style: CHAR }, { text: " and drop that", style: CHAR }],
    };
    const doc: Document = {
      section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
      blocks: [para],
      bookmarks: { mark: { start: { blockId: "p0", offset: 0 }, end: { blockId: "p0", offset: 9 } } },
    };
    expect(coveredText(doc, "mark")).toBe("keep this");

    const r1 = await roundtrip(doc);
    // Runs coalesced into one on import ("keep this and drop that").
    const p1 = r1.blocks[0] as Paragraph;
    expect(p1.runs.length).toBe(1);
    expect(coveredText(r1, "mark")).toBe("keep this");

    const r2 = await roundtrip(r1);
    expect(coveredText(r2, "mark")).toBe("keep this"); // did not grow to the run end
  }, 60_000);

  it("keeps a bookmark boundary on the correct side of an adjacent footnote reference", async () => {
    // A footnote reference is 0-width in the import IR but paints its number ("1")
    // in the model. A bookmark that STARTS immediately after the reference must
    // begin after that number; one that ENDS immediately before it must stop
    // before it. Offsets alone can't tell the two apart (both share the ref's
    // offset) — XML order does. Here: "note<fnRef>[start]body[end]", so the
    // bookmark must cover exactly "body", not "1body".
    const CHAR: CharStyle = {
      fontFamily: "Times New Roman", fontSizePx: 16, bold: false, italic: false,
      underline: false, strikethrough: false, color: "#000",
    };
    const PARA: ParaStyle = {
      align: "left", lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 8,
      indentFirstLinePx: 0, indentLeftPx: 0,
    };
    // Model: "note" + footnote ref "1" + "body". Bookmark covers "body" (the ref's
    // number sits at offset 4, so "body" is offsets 5..9).
    const doc: Document = {
      section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
      blocks: [{
        kind: "paragraph", id: "p0", revision: 0, style: PARA,
        runs: [
          { text: "note", style: CHAR },
          { text: "1", style: { ...CHAR, footnoteRef: "fn1", verticalAlign: "super" } },
          { text: "body", style: CHAR },
        ],
      }],
      footnotes: { fn1: [{ kind: "paragraph", id: "fnp", revision: 0, style: PARA, runs: [{ text: "the note", style: CHAR }] }] },
      bookmarks: { mark: { start: { blockId: "p0", offset: 5 }, end: { blockId: "p0", offset: 9 } } },
    };
    expect(coveredText(doc, "mark")).toBe("body");

    const r1 = await roundtrip(doc);
    expect(coveredText(r1, "mark")).toBe("body"); // start not dragged before the "1"

    const r2 = await roundtrip(r1);
    expect(coveredText(r2, "mark")).toBe("body");
  }, 60_000);
});
