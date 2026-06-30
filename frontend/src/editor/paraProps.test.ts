import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, SectionProps } from "@cw/shared";
import { applyOp } from "@cw/shared";
import { setParaProps } from "./commands";
import type { Command, EditorState } from "./state";

// Exercises the command the Paragraph dialog dispatches: setParaProps applying the
// paragraph-level fields (borders, shading, line-spacing rule, widow control,
// vertical alignment, mirror indents, …) over the selection, and undoing cleanly.

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 14, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA: ParaStyle = { align: "left", lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

let n = 0;
const p = (text: string): Paragraph => ({
  kind: "paragraph", id: `p${n++}`, revision: 0,
  runs: [{ text, style: { ...CHAR } }],
  style: { ...PARA },
});
const docOf = (...ps: Paragraph[]): Document => ({ section: SECTION, blocks: ps });

interface RunResult { doc: Document; inverses: ReturnType<typeof applyOp>["inverse"][] }
const run = (state: EditorState, cmd: Command): RunResult => {
  const trn = cmd(state);
  if (!trn) throw new Error("command produced no transaction");
  let doc = state.doc;
  const inverses: RunResult["inverses"] = [];
  for (const op of trn.ops) { const r = applyOp(doc, op); doc = r.doc; inverses.push(r.inverse); }
  return { doc, inverses };
};
const undo = ({ doc, inverses }: RunResult): Document => {
  let d = doc;
  for (let i = inverses.length - 1; i >= 0; i--) d = applyOp(d, inverses[i]!).doc;
  return d;
};
const sel = (b: Paragraph) => ({ anchor: { blockId: b.id, offset: 0 }, focus: { blockId: b.id, offset: 1 } });
const para = (d: Document, id: string): Paragraph => d.blocks.find((b) => b.id === id) as Paragraph;

describe("setParaProps — Paragraph dialog fields", () => {
  it("applies borders + shading + flags and restores them on undo", () => {
    const a = p("hello");
    const state: EditorState = { doc: docOf(a), selection: sel(a) };
    const patch: Partial<ParaStyle> = {
      borders: { top: { color: "#e8710a", widthPx: 3, style: "double" }, left: { color: "#e8710a", widthPx: 3, style: "double" } },
      shading: "#eef4ff",
      contextualSpacing: true,
      widowControl: false,
      mirrorIndents: true,
      suppressLineNumbers: true,
      adjustRightInd: true,
      textAlignment: "center",
    };
    const res = run(state, setParaProps(patch));
    const styled = para(res.doc, a.id).style;
    expect(styled.borders).toEqual(patch.borders);
    expect(styled.shading).toBe("#eef4ff");
    expect(styled.contextualSpacing).toBe(true);
    expect(styled.widowControl).toBe(false);
    expect(styled.mirrorIndents).toBe(true);
    expect(styled.suppressLineNumbers).toBe(true);
    expect(styled.adjustRightInd).toBe(true);
    expect(styled.textAlignment).toBe("center");

    const back = undo(res);
    const restored = para(back, a.id).style;
    expect(restored.borders).toBeUndefined();
    expect(restored.shading).toBeUndefined();
    expect(restored.contextualSpacing).toBeUndefined();
    expect(restored.widowControl).toBeUndefined();
    expect(restored.textAlignment).toBeUndefined();
  });

  it("sets a fixed line-spacing rule (exact) and undoes back to the multiplier", () => {
    const a = p("spacing");
    const state: EditorState = { doc: docOf(a), selection: sel(a) };
    const res = run(state, setParaProps({ lineRule: "exact", lineHeightPx: 28 }));
    const styled = para(res.doc, a.id).style;
    expect(styled.lineRule).toBe("exact");
    expect(styled.lineHeightPx).toBe(28);
    expect(styled.lineHeight).toBe(1.5); // multiplier untouched

    const restored = para(undo(res), a.id).style;
    expect(restored.lineRule).toBeUndefined();
    expect(restored.lineHeightPx).toBeUndefined();
  });

  it("multiple-rule reset clears any prior fixed line height", () => {
    const a: Paragraph = { ...p("fixed"), style: { ...PARA, lineRule: "atLeast", lineHeightPx: 24 } };
    const state: EditorState = { doc: docOf(a), selection: sel(a) };
    const resetPatch: Partial<ParaStyle> = { lineHeight: 2 };
    Object.assign(resetPatch, { lineRule: undefined, lineHeightPx: undefined });
    const res = run(state, setParaProps(resetPatch));
    const styled = para(res.doc, a.id).style;
    expect(styled.lineRule).toBeUndefined();
    expect(styled.lineHeightPx).toBeUndefined();
    expect(styled.lineHeight).toBe(2);

    const restored = para(undo(res), a.id).style;
    expect(restored.lineRule).toBe("atLeast");
    expect(restored.lineHeightPx).toBe(24);
  });

  it("applies across a multi-paragraph selection", () => {
    const a = p("first");
    const b = p("second");
    const state: EditorState = { doc: docOf(a, b), selection: { anchor: { blockId: a.id, offset: 0 }, focus: { blockId: b.id, offset: 3 } } };
    const res = run(state, setParaProps({ shading: "#fff3e0" }));
    expect(para(res.doc, a.id).style.shading).toBe("#fff3e0");
    expect(para(res.doc, b.id).style.shading).toBe("#fff3e0");
  });
});
