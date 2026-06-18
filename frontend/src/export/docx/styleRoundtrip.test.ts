// End-to-end: every style kind a user can author must survive the real flow —
// create a style, APPLY it to content, save (export .docx), reopen (import). The
// importer prunes to USED styles, so this asserts the applied case (the one that
// matters); unapplied/never-referenced styles are intentionally not retained.
import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, SectionProps, TableBlock, TableStyle } from "@cw/shared";
import { numberListDefinition, styleType } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import { writeDocx } from "./writeDocx";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#202124" };
const PARA = { align: "left" as const, lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

const tblStyle: TableStyle = {
  id: "MyGrid", name: "My Grid", rowBandSize: 1,
  conds: {
    wholeTable: { borders: { top: { color: "#000000", widthPx: 1 }, bottom: { color: "#000000", widthPx: 1 }, left: { color: "#000000", widthPx: 1 }, right: { color: "#000000", widthPx: 1 } } },
    firstRow: { shading: "#2e7d32", char: { bold: true, color: "#ffffff" } },
    band1Horz: { shading: "#e8f5e9" },
  },
};

// A document that USES one of each style kind, as the editor would after the user
// creates + applies them.
function authoredDoc(): Document {
  const headingPara: Paragraph = {
    kind: "paragraph", id: "h1", revision: 0,
    style: { ...PARA, align: "center", spaceBeforePx: 18, namedStyle: "MyHeading" },
    runs: [{ text: "Styled Heading", style: { ...CHAR, bold: true, color: "#114488", fontSizePx: 28 } }],
  };
  const emphPara: Paragraph = {
    kind: "paragraph", id: "p2", revision: 0, style: { ...PARA },
    runs: [
      { text: "Plain then ", style: { ...CHAR } },
      { text: "emphasized", style: { ...CHAR, italic: true, color: "#cc0000", charStyleId: "MyEmph" } },
      { text: " text.", style: { ...CHAR } },
    ],
  };
  const listPara: Paragraph = {
    kind: "paragraph", id: "li1", revision: 0,
    style: { ...PARA, list: { listId: "mylist", level: 0 } },
    runs: [{ text: "First numbered item", style: { ...CHAR } }],
  };
  const table: TableBlock = {
    kind: "table", id: "T1", revision: 0, colFractions: [0.5, 0.5],
    styleId: "MyGrid", condOverrides: { firstRow: true, bandRows: true },
    rows: [
      { cells: [{ id: "c1", blocks: [{ kind: "paragraph", id: "tc1", revision: 0, style: { ...PARA }, runs: [{ text: "H1", style: { ...CHAR } }] }], shading: "#2e7d32" }, { id: "c2", blocks: [{ kind: "paragraph", id: "tc2", revision: 0, style: { ...PARA }, runs: [{ text: "H2", style: { ...CHAR } }] }], shading: "#2e7d32" }] },
      { cells: [{ id: "c3", blocks: [{ kind: "paragraph", id: "tc3", revision: 0, style: { ...PARA }, runs: [{ text: "a", style: { ...CHAR } }] }] }, { id: "c4", blocks: [{ kind: "paragraph", id: "tc4", revision: 0, style: { ...PARA }, runs: [{ text: "b", style: { ...CHAR } }] }] }] },
    ],
  };
  return {
    section: SECTION,
    stylesheet: {
      defaultStyleId: "Normal",
      styles: [
        { id: "Normal", name: "Normal", type: "paragraph", char: { ...CHAR }, para: { ...PARA } },
        { id: "MyHeading", name: "My Heading", type: "paragraph", basedOn: "Normal", char: { bold: true, color: "#114488", fontSizePx: 28 }, para: { align: "center", spaceBeforePx: 18 } },
        { id: "MyEmph", name: "My Emphasis", type: "character", char: { italic: true, color: "#cc0000" }, para: {} },
      ],
    },
    lists: { mylist: numberListDefinition("mylist", "decimal", ".") },
    tableStyles: { MyGrid: tblStyle },
    blocks: [headingPara, emphPara, listPara, table],
  };
}

describe("all style kinds survive create→apply→save→reopen", () => {
  const back = runImport(writeDocx(authoredDoc()).bytes).doc;

  it("paragraph style", () => {
    const s = back.stylesheet!.styles.find((x) => x.id === "MyHeading");
    expect(s).toBeTruthy();
    expect(styleType(s!)).toBe("paragraph");
    expect(s!.char.bold).toBe(true);
    expect(s!.char.color).toBe("#114488");
    expect(s!.para.align).toBe("center");
    expect((back.blocks[0] as Paragraph).style.namedStyle).toBe("MyHeading"); // reference kept
  });

  it("character style", () => {
    const s = back.stylesheet!.styles.find((x) => x.id === "MyEmph");
    expect(s).toBeTruthy();
    expect(styleType(s!)).toBe("character");
    expect(s!.char.italic).toBe(true);
    const emph = (back.blocks[1] as Paragraph).runs.find((r) => r.text === "emphasized");
    expect(emph?.style.charStyleId).toBe("MyEmph"); // reference kept
  });

  it("list style", () => {
    const defs = Object.values(back.lists ?? {});
    expect(defs.length).toBe(1);
    expect(defs[0]!.levels[0]!.format).toBe("decimal");
    // The list id is remapped to its numId on export; the paragraph references the
    // same key, so membership stays intact.
    const listPara = back.blocks[2] as Paragraph;
    expect(listPara.style.list).toBeTruthy();
    expect(back.lists![listPara.style.list!.listId]).toBeTruthy();
  });

  it("table style", () => {
    const s = back.tableStyles?.MyGrid;
    expect(s).toBeTruthy();
    expect(s!.conds.firstRow?.shading).toBe("#2e7d32");
    expect(s!.conds.band1Horz?.shading).toBe("#e8f5e9");
    const t = back.blocks.find((b): b is TableBlock => b.kind === "table")!;
    expect(t.styleId).toBe("MyGrid");
    expect(t.condOverrides?.firstRow).toBe(true);
  });
});

// Unapplied styles must ALSO survive: the importer now preserves every defined
// style (paragraph/character/list/table), so a style authored in the editor
// round-trips even before it's applied to any content (matches Word).
describe("unapplied styles survive reopen", () => {
  const doc: Document = {
    section: SECTION,
    stylesheet: {
      defaultStyleId: "Normal",
      styles: [
        { id: "Normal", name: "Normal", type: "paragraph", char: { ...CHAR }, para: { ...PARA } },
        { id: "UnusedPara", name: "Unused Para", type: "paragraph", basedOn: "Normal", char: { bold: true }, para: {} },
        { id: "UnusedChar", name: "Unused Char", type: "character", char: { italic: true }, para: {} },
      ],
    },
    lists: { unusedList: numberListDefinition("unusedList", "lowerLetter", ")") },
    tableStyles: { UnusedTable: { ...tblStyle, id: "UnusedTable", name: "Unused Table" } },
    blocks: [{ kind: "paragraph", id: "p", revision: 0, style: { ...PARA }, runs: [{ text: "no styles used here", style: { ...CHAR } }] }],
  };
  const back = runImport(writeDocx(doc).bytes).doc;

  it("keeps the unapplied paragraph/character/list/table styles", () => {
    expect(back.stylesheet?.styles.find((s) => s.id === "UnusedPara")?.char.bold).toBe(true);
    const uc = back.stylesheet?.styles.find((s) => s.id === "UnusedChar");
    expect(styleType(uc!)).toBe("character");
    expect(Object.values(back.lists ?? {}).some((l) => l.levels[0]!.format === "lowerLetter")).toBe(true);
    expect(back.tableStyles?.UnusedTable).toBeTruthy();
  });
});
