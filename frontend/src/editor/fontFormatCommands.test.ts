import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, SectionProps } from "@cw/shared";
import { applyOp } from "@cw/shared";
import { setCharStyle, toggleCharStyle } from "./commands";
import type { Command, EditorState } from "./state";

// Commands behind issue #84's Font dialog: the caps / small-caps / double-strike
// quick toggles (toggleCharStyle) and the dialog's Apply patch (setCharStyle).
// Both restyle the selected runs and must round-trip cleanly through undo.

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 14, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA: ParaStyle = { align: "left", lineHeight: 1.35, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
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
const range = (b: Paragraph, from: number, to: number) => ({ anchor: { blockId: b.id, offset: from }, focus: { blockId: b.id, offset: to } });
const firstRun = (doc: Document) => (doc.blocks[0] as Paragraph).runs[0]!;

describe("toggleCharStyle — caps / smallCaps / doubleStrikethrough", () => {
  it("applies caps over the selection and undo restores it", () => {
    const a = p("hello");
    const state: EditorState = { doc: docOf(a), selection: range(a, 0, 5) };
    const res = run(state, toggleCharStyle("caps"));
    expect(firstRun(res.doc).style.caps).toBe(true);
    const back = undo(res);
    expect(firstRun(back).style.caps).toBeUndefined();
  });

  it("toggles smallCaps and doubleStrikethrough off again when already set", () => {
    const a = p("hi");
    a.runs = [{ text: "hi", style: { ...CHAR, smallCaps: true, doubleStrikethrough: true } }];
    const state: EditorState = { doc: docOf(a), selection: range(a, 0, 2) };
    expect(firstRun(run(state, toggleCharStyle("smallCaps")).doc).style.smallCaps).toBe(false);
    expect(firstRun(run(state, toggleCharStyle("doubleStrikethrough")).doc).style.doubleStrikethrough).toBe(false);
  });

  it("is a no-op on a collapsed caret", () => {
    const a = p("x");
    const collapsed: EditorState = { doc: docOf(a), selection: range(a, 0, 0) };
    expect(toggleCharStyle("caps")(collapsed)).toBeNull();
  });
});

describe("setCharStyle — Font dialog patch (apply + undo)", () => {
  it("bakes underline style/color, position, scaling, spacing, kerning, emphasis and effects", () => {
    const a = p("styled");
    const state: EditorState = { doc: docOf(a), selection: range(a, 0, 6) };
    const patch: Partial<CharStyle> = {
      underline: true,
      underlineStyle: "wave",
      underlineColor: "#ff0000",
      positionPx: 3,
      widthScalePct: 150,
      letterSpacingPx: 1.5,
      kerningMinPx: 8,
      emphasisMark: "dot",
      outline: true,
      shadow: true,
      fitTextPx: 120,
    };
    const res = run(state, setCharStyle(patch));
    const s = firstRun(res.doc).style;
    expect(s.underline).toBe(true);
    expect(s.underlineStyle).toBe("wave");
    expect(s.underlineColor).toBe("#ff0000");
    expect(s.positionPx).toBe(3);
    expect(s.widthScalePct).toBe(150);
    expect(s.letterSpacingPx).toBe(1.5);
    expect(s.kerningMinPx).toBe(8);
    expect(s.emphasisMark).toBe("dot");
    expect(s.outline).toBe(true);
    expect(s.shadow).toBe(true);
    expect(s.fitTextPx).toBe(120);

    const back = undo(res);
    const b = firstRun(back).style;
    expect(b.underline).toBe(false);
    expect(b.underlineStyle).toBeUndefined();
    expect(b.positionPx).toBeUndefined();
    expect(b.emphasisMark).toBeUndefined();
    expect(b.outline).toBeUndefined();
    expect(b.fitTextPx).toBeUndefined();
  });

  it("turns boolean effects off by patching explicit false (dialog's clear path)", () => {
    const a = p("E");
    a.runs = [{ text: "E", style: { ...CHAR, caps: true, smallCaps: true, outline: true } }];
    const state: EditorState = { doc: docOf(a), selection: range(a, 0, 1) };
    const s = firstRun(run(state, setCharStyle({ caps: false, smallCaps: false, outline: false })).doc).style;
    expect(s.caps).toBe(false);
    expect(s.smallCaps).toBe(false);
    expect(s.outline).toBe(false);
  });

  it("clears underline style/color by patching undefined", () => {
    const a = p("u");
    a.runs = [{ text: "u", style: { ...CHAR, underline: true, underlineStyle: "double", underlineColor: "#00ff00" } }];
    const state: EditorState = { doc: docOf(a), selection: range(a, 0, 1) };
    const res = run(state, setCharStyle({ underline: false, underlineStyle: undefined, underlineColor: undefined }));
    const s = firstRun(res.doc).style;
    expect(s.underline).toBe(false);
    expect(s.underlineStyle).toBeUndefined();
    expect(s.underlineColor).toBeUndefined();
  });
});
