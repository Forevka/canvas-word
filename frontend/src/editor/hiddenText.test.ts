import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, SectionProps } from "@cw/shared";
import { applyOp } from "@cw/shared";
import { deleteBackward } from "./commands";
import type { EditorState } from "./state";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 14, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA: ParaStyle = { align: "left", lineHeight: 1.35, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

let n = 0;
const vis = (text: string): Paragraph => ({ kind: "paragraph", id: `v${n++}`, revision: 0, runs: [{ text, style: CHAR }], style: PARA });
const hidden = (text: string): Paragraph => ({ kind: "paragraph", id: `h${n++}`, revision: 0, runs: [{ text, style: { ...CHAR, hidden: true } }], style: PARA });

const run = (state: EditorState): Document => {
  const trn = deleteBackward()(state);
  if (!trn) throw new Error("no transaction");
  let doc = state.doc;
  for (const op of trn.ops) doc = applyOp(doc, op).doc;
  return doc;
};

describe("hidden text — deletion protection", () => {
  it("survives Select-All → Delete spanning it", () => {
    const a = vis("AAA");
    const anchor = hidden("secret");
    const b = vis("BBB");
    const doc: Document = { section: SECTION, blocks: [a, anchor, b] };
    // Select-all over the VISIBLE range (start of first visible → end of last visible).
    const state: EditorState = {
      doc,
      selection: { anchor: { blockId: a.id, offset: 0 }, focus: { blockId: b.id, offset: 3 } },
    };
    const out = run(state);
    const survivor = out.blocks.find((bl) => bl.kind === "paragraph" && bl.runs.some((r) => r.text === "secret"));
    expect(survivor).toBeDefined(); // hidden anchor preserved
    expect((survivor as Paragraph).runs[0]!.style.hidden).toBe(true);
    // visible text is gone
    const visibleText = out.blocks
      .filter((bl): bl is Paragraph => bl.kind === "paragraph")
      .flatMap((p) => p.runs.filter((r) => !r.style.hidden).map((r) => r.text))
      .join("");
    expect(visibleText).toBe("");
  });
});
