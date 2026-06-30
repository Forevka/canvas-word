import { describe, expect, it } from "vitest";
import type { Document, Paragraph, SectionProps } from "@cw/shared";
import { applyOp } from "@cw/shared";
import { applyPageSetup, insertSectionBreak, pageSetupAt } from "./commands";
import type { Command, EditorState } from "./state";

const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

let n = 0;
const para = (text: string): Paragraph => ({
  kind: "paragraph",
  id: `p${n++}`,
  revision: 0,
  runs: [{ text, style: { fontFamily: "Georgia, serif", fontSizePx: 14, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" } }],
  style: { align: "left", lineHeight: 1.35, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 },
});

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

describe("applyPageSetup — section start + line numbering on the final section", () => {
  it("seeds defaults from doc.section", () => {
    const a = para("hello");
    const geo = pageSetupAt({ doc: { section: SECTION, blocks: [a] }, selection: { anchor: { blockId: a.id, offset: 0 }, focus: { blockId: a.id, offset: 0 } } });
    expect(geo.breakType).toBeNull();
    expect(geo.lineNumbering).toBeNull();
  });

  it("writes breakType + lineNumbering to doc.section and undoes cleanly", () => {
    const a = para("body");
    const state: EditorState = { doc: { section: SECTION, blocks: [a] }, selection: { anchor: { blockId: a.id, offset: 0 }, focus: { blockId: a.id, offset: 0 } } };
    const geo = { ...pageSetupAt(state), breakType: "evenPage" as const, lineNumbering: { countBy: 2, restart: "newSection" as const, distancePx: 24 } };
    const res = run(state, applyPageSetup(geo));
    expect(res.doc.section.breakType).toBe("evenPage");
    expect(res.doc.section.lineNumbering).toEqual({ countBy: 2, restart: "newSection", distancePx: 24 });
    // undo restores the bare section
    const back = undo(res);
    expect(back.section.breakType).toBeUndefined();
    expect(back.section.lineNumbering).toBeUndefined();
  });

  it("clearing breakType/lineNumbering (null geometry) removes them", () => {
    const a = para("body");
    const seeded: SectionProps = { ...SECTION, breakType: "oddPage", lineNumbering: { countBy: 5 } };
    const state: EditorState = { doc: { section: seeded, blocks: [a] }, selection: { anchor: { blockId: a.id, offset: 0 }, focus: { blockId: a.id, offset: 0 } } };
    const geo = pageSetupAt(state);
    expect(geo.breakType).toBe("oddPage");
    expect(geo.lineNumbering).toEqual({ countBy: 5 });
    const res = run(state, applyPageSetup({ ...geo, breakType: null, lineNumbering: null }));
    expect(res.doc.section.breakType).toBeUndefined();
    expect(res.doc.section.lineNumbering).toBeUndefined();
    const back = undo(res);
    expect(back.section.breakType).toBe("oddPage");
    expect(back.section.lineNumbering).toEqual({ countBy: 5 });
  });
});

describe("applyPageSetup — section start + line numbering on a mid-document break", () => {
  it("patches the section-break paragraph's type + props.lineNumbering and undoes", () => {
    const head = para("section one");
    head.style.sectionBreak = { type: "nextPage", props: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } } };
    const tail = para("section two");
    const state: EditorState = { doc: { section: SECTION, blocks: [head, tail] }, selection: { anchor: { blockId: head.id, offset: 0 }, focus: { blockId: head.id, offset: 0 } } };

    const geo = pageSetupAt(state);
    expect(geo.breakType).toBe("nextPage");
    expect(geo.lineNumbering).toBeNull();

    const res = run(state, applyPageSetup({ ...geo, breakType: "oddPage", lineNumbering: { countBy: 3 } }));
    const brk = (res.doc.blocks[0] as Paragraph).style.sectionBreak!;
    expect(brk.type).toBe("oddPage");
    expect(brk.props.lineNumbering).toEqual({ countBy: 3 });

    const back = undo(res);
    const brk0 = (back.blocks[0] as Paragraph).style.sectionBreak!;
    expect(brk0.type).toBe("nextPage");
    expect(brk0.props.lineNumbering).toBeUndefined();
  });

  it("insertSectionBreak carries the section's line numbering onto the new break", () => {
    const a = para("body");
    const lined: SectionProps = { ...SECTION, lineNumbering: { countBy: 2 } };
    const state: EditorState = { doc: { section: lined, blocks: [a] }, selection: { anchor: { blockId: a.id, offset: 1 }, focus: { blockId: a.id, offset: 1 } } };
    const res = run(state, insertSectionBreak());
    const brk = (res.doc.blocks[0] as Paragraph).style.sectionBreak!;
    expect(brk.type).toBe("nextPage");
    expect(brk.props.lineNumbering).toEqual({ countBy: 2 });
  });
});
