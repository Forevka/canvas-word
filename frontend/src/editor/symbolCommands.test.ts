import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, SectionProps } from "@cw/shared";
import { applyOp } from "@cw/shared";
import { insertSymbolCmd } from "./commands";
import type { Command, EditorState } from "./state";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 14, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA: ParaStyle = { align: "left", lineHeight: 1.35, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

const para = (text: string): Paragraph => ({ kind: "paragraph", id: "p1", revision: 0, runs: [{ text, style: { ...CHAR } }], style: { ...PARA } });
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
const para0 = (doc: Document): Paragraph => doc.blocks.find((b): b is Paragraph => b.kind === "paragraph")!;

describe("insertSymbolCmd", () => {
  const caretEnd = (doc: Document) => {
    const pid = para0(doc).id;
    return { anchor: { blockId: pid, offset: 5 }, focus: { blockId: pid, offset: 5 } };
  };

  it("inserts a w:sym run carrying the marker, decoded glyph, and symbol font", () => {
    const doc = docOf(para("Hello"));
    const res = run({ doc, selection: caretEnd(doc) }, insertSymbolCmd("Wingdings", "F0E0"));
    const sym = para0(res.doc).runs.find((r) => r.style.symbol);
    expect(sym).toBeDefined();
    expect(sym!.style.symbol).toEqual({ font: "Wingdings", char: "F0E0" });
    expect(sym!.text).toBe(String.fromCodePoint(0xf0e0)); // decoded glyph, not the hex
    expect(sym!.style.fontFamily).toBe("Wingdings, sans-serif");
  });

  it("upper-cases the stored hex code point and inherits the surrounding style", () => {
    const doc = docOf(para("Hello"));
    const res = run({ doc, selection: caretEnd(doc) }, insertSymbolCmd("Symbol", "f0b7"));
    const sym = para0(res.doc).runs.find((r) => r.style.symbol)!;
    expect(sym.style.symbol!.char).toBe("F0B7");
    expect(sym.style.fontSizePx).toBe(14); // inherited from the caret's run
  });

  it("undo removes the inserted glyph and restores the original text", () => {
    const doc = docOf(para("Hello"));
    const res = run({ doc, selection: caretEnd(doc) }, insertSymbolCmd("Wingdings", "F04A"));
    expect(para0(res.doc).runs.some((r) => r.style.symbol)).toBe(true);
    const back = undo(res);
    expect(para0(back).runs.some((r) => r.style.symbol)).toBe(false);
    expect(para0(back).runs.map((r) => r.text).join("")).toBe("Hello");
  });

  it("returns null when there is no selection (nothing to insert into)", () => {
    const doc = docOf(para("Hello"));
    expect(insertSymbolCmd("Wingdings", "F0E0")({ doc, selection: null })).toBeNull();
  });
});
