// Pagination stability: a doc exported to .docx and re-imported must lay out to
// the SAME page count. Guards the common (text) path against rounding/encoding
// regressions in the import↔export unit conversions.
//
// NOTE: complex real-world documents (many tables, mixed fonts) can still drift by
// a few pages on round-trip — a separate, deeper layout-fidelity issue (table
// column-width quantization, measurement) tracked independently. It does NOT affect
// the recalc endpoint, which patches the original bytes in place and never re-exports.

import { beforeAll, describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, SectionProps } from "@cw/shared";
import { runImport } from "../import/docx/pipeline";
import { runExport } from "./pipeline";
import { installMeasureHost } from "./shared/measureHost";
import { createLayoutEngine } from "../layout/engine";

const SECTION: SectionProps = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
};
const CHAR: CharStyle = {
  fontFamily: "Times New Roman",
  fontSizePx: 16,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#000",
};
const PARA: ParaStyle = {
  align: "left",
  lineHeight: 1.5,
  spaceBeforePx: 0,
  spaceAfterPx: 8,
  indentFirstLinePx: 0,
  indentLeftPx: 0,
};

const pageCount = (doc: Document): number => createLayoutEngine().layout(doc).pages.length;

describe("docx round-trip pagination stability (text)", () => {
  beforeAll(async () => {
    await installMeasureHost();
  });

  it("preserves page count across export → re-import", async () => {
    const sentence = "The quick brown fox jumps over the lazy dog and keeps on running across the wide open field. ";
    const blocks: Paragraph[] = Array.from({ length: 200 }, (_, i) => ({
      kind: "paragraph",
      id: `p${i}`,
      revision: 0,
      runs: [{ text: sentence.repeat(2) + `(${i})`, style: CHAR }],
      style: PARA,
    }));
    const doc: Document = { section: SECTION, blocks };

    const before = pageCount(doc);
    expect(before).toBeGreaterThan(3); // genuinely multi-page

    const { bytes } = await runExport(doc, "docx");
    const reimported = runImport(bytes).doc;
    expect(pageCount(reimported)).toBe(before);
  });
});
