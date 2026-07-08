import { describe, expect, it } from "vitest";
import type { Document, Op, Paragraph, SectionProps } from "@cw/shared";
import { applyOp } from "@cw/shared";
import {
  computePageGroups,
  deletePageGroupOps,
  groupStartIndices,
  pinPageBreakBeforeOps,
  reorderPageGroupsOps,
} from "./pageGroups";
import type { LayoutTree } from "./layoutTree";

const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

const RUN_STYLE = { fontFamily: "Georgia, serif", fontSizePx: 14, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const BASE_PARA = { align: "left" as const, lineHeight: 1.35, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };

interface Opts { pageBreak?: boolean; sectionBreak?: boolean }
const para = (id: string, o: Opts = {}): Paragraph => ({
  kind: "paragraph",
  id,
  revision: 0,
  runs: [{ text: id, style: { ...RUN_STYLE } }],
  style: {
    ...BASE_PARA,
    ...(o.pageBreak ? { pageBreakBefore: true } : {}),
    ...(o.sectionBreak ? { sectionBreak: { type: "nextPage" as const, props: {} } } : {}),
  },
});

const docOf = (...blocks: Paragraph[]): Document => ({ section: SECTION, blocks });
const ids = (doc: Document): string[] => doc.blocks.map((b) => b.id);
const paraById = (doc: Document, id: string): Paragraph => doc.blocks.find((b) => b.id === id) as Paragraph;

const applyAll = (doc: Document, ops: Op[]): { doc: Document; inverses: Op[] } => {
  let d = doc;
  const inverses: Op[] = [];
  for (const op of ops) { const r = applyOp(d, op); d = r.doc; inverses.push(r.inverse); }
  return { doc: d, inverses };
};
const undoAll = (doc: Document, inverses: Op[]): Document => {
  let d = doc;
  for (let i = inverses.length - 1; i >= 0; i--) d = applyOp(d, inverses[i]!).doc;
  return d;
};

describe("computePageGroups — boundaries", () => {
  it("a document with no hard breaks is one group", () => {
    const doc = docOf(para("A"), para("B"), para("C"));
    expect(groupStartIndices(doc)).toEqual([0]);
    const groups = computePageGroups(doc);
    expect(groups).toHaveLength(1);
    expect(groups[0]!.blockIds).toEqual(["A", "B", "C"]);
  });

  it("pageBreakBefore starts a new group", () => {
    const doc = docOf(para("A"), para("B", { pageBreak: true }), para("C", { pageBreak: true }));
    expect(groupStartIndices(doc)).toEqual([0, 1, 2]);
    expect(computePageGroups(doc).map((g) => g.blockIds)).toEqual([["A"], ["B"], ["C"]]);
  });

  it("a section break terminates a group (break paragraph is the group's LAST block)", () => {
    const doc = docOf(para("A"), para("B", { sectionBreak: true }), para("C"), para("D", { pageBreak: true }));
    expect(groupStartIndices(doc)).toEqual([0, 2, 3]);
    expect(computePageGroups(doc).map((g) => g.blockIds)).toEqual([["A", "B"], ["C"], ["D"]]);
  });
});

describe("computePageGroups — visual-page units (tree-aware)", () => {
  // Minimal LayoutTree — computePageGroups only reads page.index and each page's
  // first placed block id (page.blocks[0].blockId).
  const fakeTree = (pages: string[][]): LayoutTree =>
    ({ pages: pages.map((blockIds, i) => ({ index: i, blocks: blockIds.map((id) => ({ blockId: id })) })) }) as unknown as LayoutTree;

  it("maps each group to the page it starts", () => {
    const doc = docOf(para("A"), para("B", { pageBreak: true }), para("C", { pageBreak: true }));
    const units = computePageGroups(doc, fakeTree([["A"], ["B"], ["C"]]));
    expect(units.map((u) => u.blockIds)).toEqual([["A"], ["B"], ["C"]]);
    expect(units.map((u) => u.pageIndices)).toEqual([[0], [1], [2]]);
  });

  it("merges a group that doesn't begin a fresh page into the unit before it", () => {
    // B's page break is swallowed (e.g. a hidden separator), so A and B render on
    // the same page — they must be one movable unit.
    const doc = docOf(para("A"), para("B", { pageBreak: true }));
    const units = computePageGroups(doc, fakeTree([["A", "B"]]));
    expect(units).toHaveLength(1);
    expect(units[0]!.blockIds).toEqual(["A", "B"]);
    expect(units[0]!.pageIndices).toEqual([0]);
    // The merge absorbs the would-be zero-page group, so no unit is left with
    // empty pageIndices (the state repairPagination would otherwise flag).
    expect(units.every((u) => u.pageIndices.length > 0)).toBe(true);
  });

  it("keeps a group that spans multiple pages as one unit", () => {
    const doc = docOf(para("A"), para("B", { pageBreak: true }));
    // A begins page 0 and continues onto page 1; B begins page 2.
    const units = computePageGroups(doc, fakeTree([["A"], ["A"], ["B"]]));
    expect(units.map((u) => u.pageIndices)).toEqual([[0, 1], [2]]);
  });
});

describe("reorderPageGroupsOps", () => {
  it("moving a group to the front re-sequences blocks and normalizes seams", () => {
    const doc = docOf(para("A"), para("B", { pageBreak: true }), para("C", { pageBreak: true }), para("D", { pageBreak: true }));
    const ops = reorderPageGroupsOps(doc, "D", "D", "A"); // move [D] before [A]
    const { doc: after, inverses } = applyAll(doc, ops);

    expect(ids(after)).toEqual(["D", "A", "B", "C"]);
    // D is now first → its leading break is stripped; A becomes a group start → gains one.
    expect(paraById(after, "D").style.pageBreakBefore).toBeFalsy();
    expect(paraById(after, "A").style.pageBreakBefore).toBe(true);
    expect(paraById(after, "B").style.pageBreakBefore).toBe(true);

    // Undo restores original order AND styles.
    const back = undoAll(after, inverses);
    expect(ids(back)).toEqual(["A", "B", "C", "D"]);
    expect(paraById(back, "A").style.pageBreakBefore).toBeFalsy();
    expect(paraById(back, "D").style.pageBreakBefore).toBe(true);
  });

  it("moving a group to the end places it last", () => {
    const doc = docOf(para("A"), para("B", { pageBreak: true }), para("C", { pageBreak: true }));
    const ops = reorderPageGroupsOps(doc, "A", "A", null); // move [A] to the end
    const { doc: after } = applyAll(doc, ops);
    expect(ids(after)).toEqual(["B", "C", "A"]);
    // B is now first (break stripped); A now needs a break to stay a new page.
    expect(paraById(after, "B").style.pageBreakBefore).toBeFalsy();
    expect(paraById(after, "A").style.pageBreakBefore).toBe(true);
  });

  it("moving a full section leaves its terminating section break in place (other sections don't reflow)", () => {
    // [A, Bsec] [T, Tc] [X, Xt, Xsec]. The TOC (T,Tc) lives in the section
    // terminated by Xsec. Moving the [X…Xsec] section before the TOC must NOT
    // drag Xsec in front of the TOC — else the TOC changes section and reflows.
    const doc = docOf(
      para("A"),
      para("Bsec", { sectionBreak: true }),
      para("T", { pageBreak: true }),
      para("Tc"),
      para("X", { pageBreak: true }),
      para("Xt"),
      para("Xsec", { sectionBreak: true }),
    );
    const ops = reorderPageGroupsOps(doc, "X", "Xsec", "T"); // move [X,Xt,Xsec] before [T,Tc]
    const { doc: after } = applyAll(doc, ops);
    // Content moved, but Xsec stayed after the TOC content — the TOC is still in
    // the section terminated by Xsec (its geometry is untouched).
    expect(ids(after)).toEqual(["A", "Bsec", "X", "Xt", "T", "Tc", "Xsec"]);
    expect(paraById(after, "Xsec").style.sectionBreak).toBeTruthy();
    expect(paraById(after, "X").style.pageBreakBefore).toBe(true);
  });

  it("keeps a page distinct when it loses its section-break predecessor", () => {
    // [A, Bsec] [C] [D(pageBreak)]. Move [C] to the front: A is no longer first
    // and C no longer follows a section break, so both need a pinned break to
    // stay distinct pages.
    const doc = docOf(para("A"), para("B", { sectionBreak: true }), para("C"), para("D", { pageBreak: true }));
    const ops = reorderPageGroupsOps(doc, "C", "C", "A");
    const { doc: after } = applyAll(doc, ops);
    expect(ids(after)).toEqual(["C", "A", "B", "D"]);
    expect(paraById(after, "A").style.pageBreakBefore).toBe(true);
    expect(paraById(after, "D").style.pageBreakBefore).toBe(true);
  });

  it("moves a multi-group unit (e.g. an executive summary whose tables spill onto the next group) as a whole", () => {
    // [A][B(pageBreak)][C(pageBreak)][D(pageBreak)]. Treat [B,C] as one visual
    // unit and move it to the end — both must travel together.
    const doc = docOf(para("A"), para("B", { pageBreak: true }), para("C", { pageBreak: true }), para("D", { pageBreak: true }));
    const ops = reorderPageGroupsOps(doc, "B", "C", null); // move [B,C] to the end
    const { doc: after } = applyAll(doc, ops);
    expect(ids(after)).toEqual(["A", "D", "B", "C"]);
  });

  it("an in-place, self-drop, or invalid move is a no-op", () => {
    const doc = docOf(para("A"), para("B", { pageBreak: true }));
    expect(reorderPageGroupsOps(doc, "A", "A", "A")).toEqual([]); // dropped onto itself
    expect(reorderPageGroupsOps(doc, "A", "A", "B")).toEqual([]); // already in that slot
    expect(reorderPageGroupsOps(doc, "nope", "nope", null)).toEqual([]);
  });
});

describe("pinPageBreakBeforeOps (layout-verified reorder repair)", () => {
  it("pins a break on a paragraph that lacks one", () => {
    const doc = docOf(para("A"), para("B"));
    const { doc: after } = applyAll(doc, pinPageBreakBeforeOps(doc, "B"));
    expect(paraById(after, "B").style.pageBreakBefore).toBe(true);
  });

  it("is a no-op when the paragraph already breaks or the block is missing", () => {
    const doc = docOf(para("A"), para("B", { pageBreak: true }));
    expect(pinPageBreakBeforeOps(doc, "B")).toEqual([]);
    expect(pinPageBreakBeforeOps(doc, "nope")).toEqual([]);
  });
});

describe("deletePageGroupOps", () => {
  it("removes the unit's blocks", () => {
    const doc = docOf(para("A"), para("B", { pageBreak: true }), para("C", { pageBreak: true }));
    const { doc: after } = applyAll(doc, deletePageGroupOps(doc, "B", "B"));
    expect(ids(after)).toEqual(["A", "C"]);
  });

  it("strips a trailing section break left as the last block (no empty trailing section)", () => {
    const doc = docOf(para("A"), para("B", { sectionBreak: true }), para("C"));
    // groups: [A,B] [C]. Delete [C] → B would end the doc carrying a section break.
    const ops = deletePageGroupOps(doc, "C", "C");
    const { doc: after } = applyAll(doc, ops);
    expect(ids(after)).toEqual(["A", "B"]);
    expect(paraById(after, "B").style.sectionBreak).toBeFalsy();
  });

  it("refuses to delete the whole document (never empties it)", () => {
    const doc = docOf(para("A"), para("B"));
    expect(deletePageGroupOps(doc, "A", "B")).toEqual([]);
  });
});
