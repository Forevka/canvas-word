// Layout-engine tests. The engine measures text through canvas; the setup import
// (FIRST, before the engine loads) installs a deterministic stub so widths are
// chars × ½ font-size — enough to assert pagination, geometry, and the structural
// rules (sections, tables, row spans, tabs, footnotes) without a real font engine.
import "./test-canvas-setup";

import { describe, it, expect } from "vitest";
import type {
  Block,
  CharStyle,
  Document,
  ImageBlock,
  Paragraph,
  ParaStyle,
  SectionProps,
  TableBlock,
  TableCell,
} from "../model/document";
import { createLayoutEngine, effectiveSection, resolveSections } from "./engine";
import { gridColumnCount, effectiveFractions } from "../model/ops";
import type { PlacedBlock } from "./layoutTree";

// --- builders -------------------------------------------------------------

const CHAR: CharStyle = {
  fontFamily: "Georgia, serif",
  fontSizePx: 16,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#000000",
};
const PARA: ParaStyle = {
  align: "left",
  lineHeight: 1,
  spaceBeforePx: 0,
  spaceAfterPx: 0,
  indentFirstLinePx: 0,
  indentLeftPx: 0,
};
const SECTION: SectionProps = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
};
// content box: 816-192 = 624 wide, 1056-192 = 864 tall; 16px chars at ½ size = 8px,
// so ~78 chars/line and (lineHeight 1 → 16px lines) ~54 lines/page.

let nextId = 0;
const fresh = (): string => `t${nextId++}`;

const para = (text: string, style: Partial<ParaStyle> = {}, char: Partial<CharStyle> = {}): Paragraph => ({
  kind: "paragraph",
  id: fresh(),
  revision: 0,
  runs: [{ text, style: { ...CHAR, ...char } }],
  style: { ...PARA, ...style },
});

const cell = (text: string, extra: Partial<TableCell> = {}): TableCell => ({
  id: fresh(),
  blocks: [para(text)],
  ...extra,
});

const table = (rows: TableCell[][], colFractions?: number[]): TableBlock => ({
  kind: "table",
  id: fresh(),
  revision: 0,
  rows: rows.map((cells) => ({ cells })),
  ...(colFractions ? { colFractions } : {}),
});

const image = (widthPx: number, heightPx: number): ImageBlock => ({
  kind: "image",
  id: fresh(),
  revision: 0,
  src: "blob:test",
  widthPx,
  heightPx,
  align: "center",
});

const doc = (blocks: Block[], section: Partial<SectionProps> = {}): Document => ({
  section: { ...SECTION, ...section },
  blocks,
});

const layout = (d: Document) => createLayoutEngine().layout(d);
const placedOf = (tree: ReturnType<typeof layout>, id: string): { page: number; pb: PlacedBlock } | null => {
  for (const pg of tree.pages) for (const pb of pg.blocks) if (pb.blockId === id) return { page: pg.index, pb };
  return null;
};

// --- basic pagination -----------------------------------------------------

describe("engine — pagination", () => {
  it("lays a short paragraph on one page", () => {
    const p = para("hello world");
    const tree = layout(doc([p]));
    expect(tree.pages).toHaveLength(1);
    expect(placedOf(tree, p.id)!.page).toBe(0);
  });

  it("page-break-before starts a new page", () => {
    const a = para("first");
    const b = para("second", { pageBreakBefore: true });
    const tree = layout(doc([a, b]));
    expect(tree.pages.length).toBeGreaterThanOrEqual(2);
    expect(placedOf(tree, a.id)!.page).toBe(0);
    expect(placedOf(tree, b.id)!.page).toBe(1);
  });

  it("flows past the page bottom onto a second page", () => {
    // ~60 single-line paragraphs at 16px exceed the ~864px content height.
    const paras = Array.from({ length: 60 }, (_, i) => para(`line ${i}`));
    const tree = layout(doc(paras));
    expect(tree.pages.length).toBeGreaterThanOrEqual(2);
    // first and last land on different pages
    expect(placedOf(tree, paras[0]!.id)!.page).toBe(0);
    expect(placedOf(tree, paras.at(-1)!.id)!.page).toBeGreaterThan(0);
  });

  it("keep-lines-together moves a splittable paragraph whole to the next page", () => {
    // Fill page 1 to near the bottom, then a multi-line keepLines paragraph that
    // wouldn't fit in the remaining space.
    const fillers = Array.from({ length: 52 }, (_, i) => para(`f${i}`));
    const longText = Array.from({ length: 6 }, () => "wwwwwwwwwwwwwwwwwwww").join(" ").repeat(8);
    const kept = para(longText, { keepLinesTogether: true });
    const tree = layout(doc([...fillers, kept]));
    const placed = placedOf(tree, kept.id)!;
    // it didn't split: it's entirely on one page (its first line is its block y)
    const onPage = tree.pages[placed.page]!.blocks.filter((b) => b.blockId === kept.id);
    expect(onPage).toHaveLength(1);
  });
});

// --- table grid sizing (pure) ---------------------------------------------

describe("engine — grid column count", () => {
  it("counts a plain grid by cells", () => {
    expect(gridColumnCount(table([[cell("a"), cell("b"), cell("c")]]))).toBe(3);
  });

  it("counts colSpan into the width", () => {
    const t = table([[cell("wide", { colSpan: 2 }), cell("c")]]);
    expect(gridColumnCount(t)).toBe(3);
  });

  it("counts a rowSpan hole that following rows shift past", () => {
    // col0 spans both rows; row1 has fewer cells but still reaches column 3.
    const t = table([
      [cell("photo", { rowSpan: 2 }), cell("a"), cell("b")],
      [cell("c"), cell("d")],
    ]);
    expect(gridColumnCount(t)).toBe(3);
  });

  it("uses colFractions only when its length matches the true grid width", () => {
    const t = table(
      [
        [cell("photo", { rowSpan: 2 }), cell("a"), cell("b")],
        [cell("c"), cell("d")],
      ],
      [0.5, 0.25, 0.25],
    );
    expect(effectiveFractions(t)).toEqual([0.5, 0.25, 0.25]);
    // wrong-length fractions are discarded for an equal split
    const bad = table([[cell("a"), cell("b")]], [0.3, 0.3, 0.4]);
    expect(effectiveFractions(bad)).toEqual([0.5, 0.5]);
  });
});

// --- row spans (geometry) -------------------------------------------------

describe("engine — vertical cell merge", () => {
  it("spans a rowSpan cell across its rows and shifts continuation cells past the hole", () => {
    const t = table(
      [
        [cell("photo", { rowSpan: 2 }), cell("r0c1")],
        [cell("r1c1")],
      ],
      [0.5, 0.5],
    );
    const tree = layout(doc([t]));
    const placed = placedOf(tree, t.id)!.pb.table!;
    const photo = placed.rows[0]!.cells[0]!;
    const r0c1 = placed.rows[0]!.cells[1]!;
    const r1c1 = placed.rows[1]!.cells[0]!;
    // the merged cell is taller than a single row…
    expect(photo.height).toBeGreaterThan(r0c1.height);
    // …and its height equals both rows together
    expect(photo.height).toBeCloseTo(placed.rows[0]!.height + placed.rows[1]!.height, 1);
    // the continuation row's single cell sits in column 1 (it shifted past col 0)
    expect(r1c1.x).toBeCloseTo(r0c1.x, 1);
    expect(r1c1.x).toBeGreaterThan(photo.x);
  });
});

// --- adjacent-atomic gap --------------------------------------------------

describe("engine — block flow gaps", () => {
  it("stacks two adjacent tables flush but gaps a table from following text", () => {
    const t1 = table([[cell("a")]]);
    const t2 = table([[cell("b")]]);
    const p = para("after");
    const tree = layout(doc([t1, t2, p]));
    const pt1 = placedOf(tree, t1.id)!.pb.table!;
    const pt2 = placedOf(tree, t2.id)!.pb.table!;
    const pp = placedOf(tree, p.id)!.pb;
    // adjacent tables: no gap (t2 starts exactly where t1 ends)
    expect(pt2.y).toBeCloseTo(pt1.y + pt1.height, 1);
    // table → paragraph: a breathing gap exists
    expect(pp.y).toBeGreaterThan(pt2.y + pt2.height + 1);
  });
});

// --- cell shading & borders ----------------------------------------------

describe("engine — cell shading and borders", () => {
  it("carries shading and borders onto the placed cell", () => {
    const red = { color: "#c00000", widthPx: 2 };
    const t = table([
      [cell("x", { shading: "#fff2cc", borders: { top: red, bottom: red } })],
    ]);
    const placed = placedOf(layout(doc([t])), t.id)!.pb.table!;
    const c = placed.rows[0]!.cells[0]!;
    expect(c.shading).toBe("#fff2cc");
    expect(c.borders?.top?.color).toBe("#c00000");
    expect(c.borders?.bottom?.widthPx).toBe(2);
  });
});

// --- lone-image cover fill -------------------------------------------------

describe("engine — image cover fill", () => {
  it("clips a lone image filling a cell taller than it fits", () => {
    const tall = "wwww wwww wwww wwww wwww wwww wwww wwww wwww wwww wwww wwww";
    const imgCell: TableCell = { id: fresh(), blocks: [image(60, 40)], rowSpan: 2 };
    const t = table(
      [
        [imgCell, cell(tall)],
        [cell(tall)],
      ],
      [0.4, 0.6],
    );
    const placed = placedOf(layout(doc([t])), t.id)!.pb.table!;
    const imgBlock = placed.rows[0]!.cells[0]!.blocks.find((b) => b.image)!;
    expect(imgBlock.image!.clip).toBeDefined();
    // covered: scaled at least to the cell width
    expect(imgBlock.image!.width).toBeGreaterThanOrEqual(60);
  });
});

// --- tab stops ------------------------------------------------------------

describe("engine — tab stops", () => {
  it("wraps a leading-tab paragraph: first line indented, the rest at the margin", () => {
    const long = "word ".repeat(60).trim();
    const p = para(`\t${long}`);
    const tree = layout(doc([p]));
    const lines = placedOf(tree, p.id)!.pb.lines;
    expect(lines.length).toBeGreaterThan(1);
    expect(lines[0]!.fragments[0]!.x).toBeGreaterThan(20); // first line at the tab stop
    expect(lines[1]!.fragments[0]!.x).toBeCloseTo(0, 0); // continuation at the margin
  });

  it("right-aligns a piece at a right tab stop with a leader", () => {
    const p = para("Left\tRight", { tabStops: [{ posPx: 300, align: "right", leader: "dot" }] });
    const line = placedOf(layout(doc([p])), p.id)!.pb.lines[0]!;
    const right = line.fragments.find((f) => f.text === "Right")!;
    // "Right" (5 chars × 8px = 40) ends at the 300px stop → starts near 260
    expect(right.x + right.width).toBeCloseTo(300, 0);
    expect(line.leaders?.length).toBeGreaterThan(0);
  });
});

// --- sections & page numbers ----------------------------------------------

describe("engine — sections", () => {
  it("effectiveSection inherits absent fields and applies present ones", () => {
    const base: SectionProps = { ...SECTION, columns: { count: 2, gapPx: 24 } };
    const inherit = effectiveSection(base, {});
    expect(inherit.pageWidthPx).toBe(SECTION.pageWidthPx);
    expect(inherit.columns).toEqual({ count: 2, gapPx: 24 });
    const override = effectiveSection(base, { columns: null, pageNumberStart: 5 });
    expect(override.columns).toBeUndefined(); // null = explicitly single-column
    expect(override.pageNumberStart).toBe(5);
  });

  it("resolveSections splits at a section-break paragraph", () => {
    const a = para("sec one", { sectionBreak: { type: "nextPage", props: {} } });
    const b = para("sec two");
    const sections = resolveSections(doc([a, b]));
    expect(sections).toHaveLength(2);
    expect(sections[0]!.endBlock).toBe(0);
  });

  it("honors pageNumberStart on the displayed Page.number", () => {
    const tree = layout(doc([para("x")], { pageNumberStart: 5 }));
    expect(tree.pages[0]!.number).toBe(5);
  });

  it("plain pages number from 1", () => {
    const tree = layout(doc([para("a", { pageBreakBefore: false }), para("b", { pageBreakBefore: true })]));
    expect(tree.pages[0]!.number).toBe(1);
    expect(tree.pages[1]!.number).toBe(2);
  });
});

// --- newspaper columns -----------------------------------------------------

describe("engine — columns", () => {
  it("fills column 1 then column 2 before a new page", () => {
    // enough lines to overflow one column but fit two
    const paras = Array.from({ length: 70 }, (_, i) => para(`c${i}`));
    const tree = layout(doc(paras, { columns: { count: 2, gapPx: 24 } }));
    // a 2-column page holds ~108 short lines, so everything fits on page 0…
    expect(placedOf(tree, paras[0]!.id)!.page).toBe(0);
    expect(placedOf(tree, paras.at(-1)!.id)!.page).toBe(0);
    // …with later paragraphs pushed into the right-hand column (greater x)
    const firstX = placedOf(tree, paras[0]!.id)!.pb.x;
    const lastX = placedOf(tree, paras.at(-1)!.id)!.pb.x;
    expect(lastX).toBeGreaterThan(firstX);
  });
});

// --- footnotes ------------------------------------------------------------

describe("engine — footnotes", () => {
  it("reserves a page-bottom area and places the referenced note", () => {
    const note = para("the footnote body");
    const refPara: Paragraph = {
      kind: "paragraph",
      id: fresh(),
      revision: 0,
      runs: [
        { text: "body text", style: { ...CHAR } },
        { text: "1", style: { ...CHAR, footnoteRef: "fn1", verticalAlign: "super" } },
      ],
      style: { ...PARA },
    };
    const d: Document = { ...doc([refPara]), footnotes: { fn1: [note] } };
    const tree = layout(d);
    const refPage = placedOf(tree, refPara.id)!.page;
    // the separator rule exists on the page carrying the ref…
    expect(tree.pages[refPage]!.footnoteRuleY).toBeGreaterThan(0);
    // …and the note body is placed below it
    const placedNote = placedOf(tree, note.id)!;
    expect(placedNote.page).toBe(refPage);
    expect(placedNote.pb.y).toBeGreaterThan(tree.pages[refPage]!.footnoteRuleY!);
  });
});

// --- table of contents ----------------------------------------------------

describe("engine — TOC page numbers", () => {
  it("decorates an entry with its target's live page number", () => {
    const target = para("CHAPTER FIVE", { pageBreakBefore: true });
    const entry = para("Chapter Five", { tocEntry: { targetId: target.id, level: 1 } });
    const tree = layout(doc([entry, target]));
    // target was pushed to page 2 by its page break
    expect(placedOf(tree, target.id)!.page).toBe(1);
    const toc = placedOf(tree, entry.id)!.pb.toc;
    expect(toc).toBeDefined();
    expect(toc!.numText).toBe("2"); // displayed number of the target's page
  });
});

// --- margin bands push the content box ------------------------------------

describe("engine — tall header pushes the body", () => {
  it("lowers contentTopPx when the header is taller than the top margin", () => {
    const header = Array.from({ length: 8 }, (_, i) => para(`header line ${i}`));
    const tree = layout(doc([para("body")], { header }));
    expect(tree.pages[0]!.contentTopPx).toBeGreaterThan(SECTION.marginPx.top);
  });
});
