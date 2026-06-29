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
  FieldDef,
  ImageBlock,
  Paragraph,
  ParaStyle,
  SectionProps,
  TableBlock,
  TableCell,
} from "@cw/shared";
import { createLayoutEngine, effectiveSection, resolveSections } from "./engine";
import { parseMathml } from "../mathml/parse";
import { gridColumnCount, effectiveFractions, defaultListDefinition } from "@cw/shared";
import { hitTestObject, hitTestSelectableObject, spaceMarkXs } from "./geometry";
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

const equationBlock = (mathml: string): Block => ({
  kind: "equation",
  id: fresh(),
  revision: 0,
  align: "center",
  equation: { ...parseMathml(mathml), display: true },
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

// --- full-document swap ---------------------------------------------------

describe("engine — full-document swap (cache reset)", () => {
  it("reset() clears stale lines so a new doc reusing block ids doesn't leak the old content", () => {
    // The docx importer re-mints block ids from i0 at revision 0 on EVERY import,
    // so two unrelated documents collide on (id, revision). The shared engine
    // keys its line cache on (id, revision, width) — without a reset between
    // documents, the second doc's paragraph hits the first doc's cached lines and
    // the two render merged. Model that collision directly.
    const engine = createLayoutEngine();
    const longPara = (): Paragraph => ({
      kind: "paragraph", id: "i0", revision: 0,
      runs: [{ text: "x".repeat(400), style: CHAR }], style: PARA,
    });
    const shortPara = (): Paragraph => ({
      kind: "paragraph", id: "i0", revision: 0,
      runs: [{ text: "hi", style: CHAR }], style: PARA,
    });

    const first = engine.layout(doc([longPara()]));
    const longLines = placedOf(first, "i0")!.pb.lines.length;
    expect(longLines).toBeGreaterThan(1);

    // Without reset the colliding id hits the cached multi-line layout (the bug).
    const leaked = engine.layout(doc([shortPara()]));
    expect(placedOf(leaked, "i0")!.pb.lines.length).toBe(longLines);

    // After reset the same short paragraph lays out fresh — a single line.
    engine.reset();
    const fixed = engine.layout(doc([shortPara()]));
    expect(placedOf(fixed, "i0")!.pb.lines.length).toBe(1);
  });
});

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

  it("keep-with-next carries a heading to its content's page across a blank spacer", () => {
    // A heading kept-with-next, a blank spacer paragraph, then a tall block that
    // can't fit the page tail. The blank line must NOT satisfy the keep — the
    // heading should travel to the block's page instead of being orphaned.
    const fillers = Array.from({ length: 51 }, (_, i) => para(`f${i}`));
    const heading = para("HEADING", { keepWithNext: true });
    const spacer = para(""); // empty line between the heading and its table
    const img = image(400, 300); // too tall to share the remaining space
    const tree = layout(doc([...fillers, heading, spacer, img]));
    const hp = placedOf(tree, heading.id)!;
    const ip = placedOf(tree, img.id)!;
    expect(hp.page).toBeGreaterThan(0); // not orphaned on the filler page
    expect(hp.page).toBe(ip.page); // travels with its content
  });
});

// --- contextual spacing (w:contextualSpacing) -----------------------------

describe("engine — contextual spacing", () => {
  // Pure line advance (no before/after spacing) — the reference each case compares to.
  const lineAdvance = (): number => {
    const plain = [para("a"), para("b")];
    const t = layout(doc(plain));
    return placedOf(t, plain[1]!.id)!.pb.y - placedOf(t, plain[0]!.id)!.pb.y;
  };

  it("collapses inter-paragraph spacing across a same-style run", () => {
    const adv = lineAdvance();
    const opts = { spaceBeforePx: 20, spaceAfterPx: 20, contextualSpacing: true };
    const a = para("a", opts), b = para("b", opts), c = para("c", opts);
    const tree = layout(doc([a, b, c]));
    const ya = placedOf(tree, a.id)!.pb.y;
    const yb = placedOf(tree, b.id)!.pb.y;
    const yc = placedOf(tree, c.id)!.pb.y;
    // Same style + contextualSpacing → all inter-paragraph spacing collapses, so
    // each paragraph advances by exactly one line height — as if unspaced.
    expect(yb - ya).toBe(adv);
    expect(yc - yb).toBe(adv);
  });

  it("keeps inter-paragraph spacing without the flag (control)", () => {
    const opts = { spaceBeforePx: 20, spaceAfterPx: 20 };
    const a = para("a", opts), b = para("b", opts);
    const tree = layout(doc([a, b]));
    const gap = placedOf(tree, b.id)!.pb.y - placedOf(tree, a.id)!.pb.y;
    // No flag: a.after (20) + b.before (20) on top of the line advance.
    expect(gap).toBe(lineAdvance() + 40);
  });

  it("does not collapse against a paragraph of a different style", () => {
    const a = para("a", { spaceAfterPx: 20, contextualSpacing: true, namedStyle: "Verse" });
    const b = para("b", { spaceBeforePx: 20, contextualSpacing: true, namedStyle: "Body" });
    const tree = layout(doc([a, b]));
    const gap = placedOf(tree, b.id)!.pb.y - placedOf(tree, a.id)!.pb.y;
    // Different namedStyle → suppression doesn't apply; both spacings remain.
    expect(gap).toBe(lineAdvance() + 40);
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

  it("clips cell content to the inner box so over-wide text can't paint past the border", () => {
    // Default Word cell margin is 7.2px left/right; the content clip must sit
    // inside those insets and span the full cell height.
    const t = table([[cell("a"), cell("b"), cell("c")]]);
    const placed = placedOf(layout(doc([t])), t.id)!.pb.table!;
    const mid = placed.rows[0]!.cells[1]!;
    const clip = mid.contentClip!;
    expect(clip).toBeDefined();
    expect(clip.x).toBeCloseTo(mid.x + 7.2, 1);
    expect(clip.x + clip.width).toBeCloseTo(mid.x + mid.width - 7.2, 1);
    expect(clip.y).toBeCloseTo(mid.y, 1);
    expect(clip.height).toBeCloseTo(mid.height, 1);
    // The clip is strictly inside the cell on both horizontal edges.
    expect(clip.x).toBeGreaterThan(mid.x);
    expect(clip.x + clip.width).toBeLessThan(mid.x + mid.width);
  });

  it("wraps cell text within the inner box when the cell paragraph has left/right indent", () => {
    // The wrap width must exclude the paragraph's own indents; otherwise the
    // text is painted shifted right but measured at the full inner width, so it
    // overflows the content box (and gets clipped) by the indent amount.
    const longText = "Seller financing or assumption of existing financing at non market terms";
    const c: TableCell = { id: fresh(), blocks: [para(longText, { indentLeftPx: 100, indentRightPx: 100 })] };
    const t = table([[c]]);
    const placed = placedOf(layout(doc([t])), t.id)!.pb.table!;
    const cell0 = placed.rows[0]!.cells[0]!;
    const pb = cell0.blocks[0]!;
    const clipRight = cell0.contentClip!.x + cell0.contentClip!.width;
    // Every painted glyph stays within the cell's inner content box (paint x is
    // block.x + frag.x, matching the renderer).
    for (const line of pb.lines) {
      for (const f of line.fragments) {
        expect(pb.x + f.x + f.width).toBeLessThanOrEqual(clipRight + 0.5);
      }
    }
    // Sanity: the text actually wrapped (more than one line), so the assertion
    // above is meaningful rather than vacuously true on a single short line.
    expect(pb.lines.length).toBeGreaterThan(1);
  });
});

describe("engine — hidden paragraphs", () => {
  it("a fully-hidden (w:vanish) paragraph takes zero space and isn't placed", () => {
    const a = para("first");
    const hiddenP: Paragraph = {
      kind: "paragraph",
      id: fresh(),
      revision: 0,
      runs: [{ text: "secret", style: { ...CHAR, hidden: true } }],
      style: { ...PARA },
    };
    const b = para("second");
    const withHidden = layout(doc([a, hiddenP, b]));
    const without = layout(doc([para("first"), para("second")]));
    expect(placedOf(withHidden, hiddenP.id)).toBeNull(); // not laid out
    // "second" lands at the SAME y whether or not the hidden anchor is present.
    const gapWith = placedOf(withHidden, b.id)!.pb.y - placedOf(withHidden, a.id)!.pb.y;
    const ctrl = without.pages[0]!.blocks;
    const gapWithout = ctrl[1]!.y - ctrl[0]!.y;
    expect(gapWith).toBeCloseTo(gapWithout, 1);
  });
});

describe("engine — lists inside table cells", () => {
  const listCell = (text: string): TableCell => ({
    id: fresh(),
    blocks: [para(text, { list: { listId: "numbers", level: 0 } })],
  });
  const cellMarkers = (tree: ReturnType<typeof layout>, tableId: string): string[] => {
    const tbl = placedOf(tree, tableId)!.pb.table!;
    const out: string[] = [];
    for (const row of tbl.rows) for (const c of row.cells) for (const b of c.blocks) if (b.marker) out.push(b.marker.text);
    return out;
  };

  it("numbers a list straight through cells in reading order (Word continuation)", () => {
    // row0: [a, b] ; row1: [c, d]. Numbering continues 1..4 across the cells.
    const t = table([
      [listCell("a"), listCell("b")],
      [listCell("c"), listCell("d")],
    ]);
    const d: Document = { ...doc([t]), lists: { numbers: defaultListDefinition("decimal") } };
    expect(cellMarkers(layout(d), t.id)).toEqual(["1.", "2.", "3.", "4."]);
  });

  it("hangs the marker inside the cell, left of the indented text", () => {
    const t = table([[listCell("item")]]);
    const d: Document = { ...doc([t]), lists: { numbers: defaultListDefinition("decimal") } };
    const tbl = placedOf(layout(d), t.id)!.pb.table!;
    const cellBox = tbl.rows[0]!.cells[0]!;
    const placedPara = cellBox.blocks[0]!;
    expect(placedPara.marker).toBeDefined();
    // marker sits left of the text origin but not outside the cell's left edge.
    expect(placedPara.marker!.x).toBeLessThan(placedPara.x);
    expect(placedPara.marker!.x).toBeGreaterThanOrEqual(cellBox.x);
  });
});

// --- cell margins (inner padding) -----------------------------------------

describe("engine — cell margins", () => {
  it("defaults to Word's 0 vertical padding (row ≈ one line) and ~7.2px horizontal", () => {
    const t = table([[cell("hi")]]);
    const placed = placedOf(layout(doc([t])), t.id)!.pb.table!;
    const row = placed.rows[0]!;
    // A single 16px line with no vertical padding — well under the old fixed
    // 12px-pad row (~28px), without pinning the exact line-height metric.
    expect(row.height).toBeLessThan(24);
    // Content is inset by the default horizontal margin (108 twips ≈ 7.2px).
    const content = row.cells[0]!.blocks[0]!;
    expect(content.x - row.cells[0]!.x).toBeCloseTo(7.2, 1);
  });

  it("honors an explicit cell.margin, adding it to the row height and content offset", () => {
    const plain = table([[cell("hi")]]);
    const padded = table([[cell("hi", { margin: { top: 10, right: 20, bottom: 10, left: 20 } })]]);
    const plainRow = placedOf(layout(doc([plain])), plain.id)!.pb.table!.rows[0]!.height;
    const padRow = placedOf(layout(doc([padded])), padded.id)!.pb.table!.rows[0]!;
    expect(padRow.height).toBeCloseTo(plainRow + 20, 0); // +10 top +10 bottom
    const content = padRow.cells[0]!.blocks[0]!;
    expect(content.x - padRow.cells[0]!.x).toBeCloseTo(20, 1);
    expect(content.y - padRow.y).toBeCloseTo(10, 1);
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

  // --- RTL / bidi tab stops (issue #6) ------------------------------------
  // The content box is 624px wide (816 − 96 − 96); test chars are 8px (½ × 16px).
  // In an RTL paragraph tab stops are measured from the RIGHT (start) edge and the
  // tab cells fill right-to-left (Word's w:bidi behavior).

  it("measures RTL tab stops from the right (start) edge", () => {
    const p = para("Left\tRight", { direction: "rtl", tabStops: [{ posPx: 300 }] });
    const line = placedOf(layout(doc([p])), p.id)!.pb.lines[0]!;
    const left = line.fragments.find((f) => f.text === "Left")!;
    const right = line.fragments.find((f) => f.text === "Right")!;
    // the pre-tab cell hugs the start (right) edge of the content box…
    expect(left.x + left.width).toBeCloseTo(624, 0);
    // …and the tabbed cell sits 300px IN from that right edge (not from the left).
    expect(right.x + right.width).toBeCloseTo(624 - 300, 0);
    // start-edge cell is visually to the RIGHT of the tabbed cell (RTL order).
    expect(left.x).toBeGreaterThan(right.x);
  });

  it("marks RTL-script tab cells as RTL fragments anchored at the right", () => {
    const heb1 = "אב";
    const heb2 = "גד";
    const p = para(`${heb1}\t${heb2}`, { direction: "rtl", tabStops: [{ posPx: 200 }] });
    const line = placedOf(layout(doc([p])), p.id)!.pb.lines[0]!;
    const f1 = line.fragments.find((f) => f.text === heb1)!;
    const f2 = line.fragments.find((f) => f.text === heb2)!;
    // Hebrew cells carry an odd (RTL) embedding level so the renderer right-anchors
    // them and caret/hit-testing take the bidi path.
    expect((f1.level ?? 0) & 1).toBe(1);
    expect((f2.level ?? 0) & 1).toBe(1);
    expect(f1.x + f1.width).toBeCloseTo(624, 0); // first cell at the right edge
    expect(f2.x + f2.width).toBeCloseTo(624 - 200, 0); // tabbed cell 200px from the right
  });

  it("mirrors the leader and right-tab stop for an RTL page-number row", () => {
    // "פרק" (3 chars = 24px) … dotted leader … "5" at a right tab 400px from the start.
    const p = para("פרק\t5", { direction: "rtl", tabStops: [{ posPx: 400, align: "right", leader: "dot" }] });
    const line = placedOf(layout(doc([p])), p.id)!.pb.lines[0]!;
    const num = line.fragments.find((f) => f.text === "5")!;
    // a right-aligned tab in RTL puts the cell's start (right) edge at the stop,
    // measured from the right edge → physical left edge at 624 − 400.
    expect(num.x).toBeCloseTo(624 - 400, 0);
    expect(line.leaders?.length).toBeGreaterThan(0);
    const ld = line.leaders![0]!;
    // the leader fills the gap on the mirrored side, ending against the start-edge cell.
    expect(ld.x2).toBeCloseTo(624 - 24, 0);
    expect(ld.x1).toBeLessThan(ld.x2);
  });

  it("mirrors a tab-arrow overlay on a tab-only RTL line (no text/leaders)", () => {
    // A "\t"-only line produces a formatting-mark arrow but no fragments or leaders;
    // it must still mirror so the arrow opens at the start (right) edge.
    const p = para("\t", { direction: "rtl" });
    const line = placedOf(layout(doc([p])), p.id)!.pb.lines[0]!;
    expect(line.tabArrows?.length).toBeGreaterThan(0);
    expect(line.tabArrows![0]!.x2).toBeCloseTo(624, 0);
  });

  it("leaves LTR tab stops left-anchored (no mirroring regression)", () => {
    const p = para("Left\tRight", { tabStops: [{ posPx: 300 }] });
    const line = placedOf(layout(doc([p])), p.id)!.pb.lines[0]!;
    const right = line.fragments.find((f) => f.text === "Right")!;
    expect(right.x).toBeCloseTo(300, 0); // positioned at the stop from the LEFT edge
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

  it("unequal columns place content at per-column x positions", () => {
    const paras = Array.from({ length: 70 }, (_, i) => para(`c${i}`));
    const tree = layout(
      doc(paras, {
        columns: { count: 2, gapPx: 24, cols: [{ widthPx: 200, spaceAfterPx: 40 }, { widthPx: 300, spaceAfterPx: 0 }] },
      }),
    );
    expect(placedOf(tree, paras[0]!.id)!.pb.x).toBeCloseTo(96, 0); // col 1 at the left margin
    expect(placedOf(tree, paras.at(-1)!.id)!.pb.x).toBeCloseTo(336, 0); // col 2 at 96 + 200 + 40
  });

  it("emits column separator x positions when sep is set", () => {
    const tree = layout(doc([para("x")], { columns: { count: 2, gapPx: 24, sep: true } }));
    const seps = tree.pages[0]!.columnSeparatorsX;
    expect(seps).toHaveLength(1);
    // equal columns: width = (624 - 24)/2 = 300; gap midpoint = 96 + 300 + 12 = 408.
    expect(seps![0]).toBeCloseTo(408, 0);
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

  it("reserves and places a note for a footnote ref inside a table cell", () => {
    const note = para("the footnote body");
    const refCell: TableCell = {
      id: fresh(),
      blocks: [
        {
          kind: "paragraph",
          id: fresh(),
          revision: 0,
          runs: [
            { text: "cell text", style: { ...CHAR } },
            { text: "1", style: { ...CHAR, footnoteRef: "fn1", verticalAlign: "super" } },
          ],
          style: { ...PARA },
        },
      ],
    };
    const t = table([[refCell, cell("other")]]);
    const d: Document = { ...doc([t]), footnotes: { fn1: [note] } };
    const tree = layout(d);
    const refPage = placedOf(tree, t.id)!.page;
    // separator rule reserved on the page carrying the table row with the ref…
    expect(tree.pages[refPage]!.footnoteRuleY).toBeGreaterThan(0);
    // …and the note body is placed at that page's bottom, below the rule.
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

describe("engine — footer band distance", () => {
  const footerBottom = (tree: ReturnType<typeof layout>): number =>
    Math.max(...tree.pages[0]!.footer!.flatMap((b) => b.lines.map((l) => b.y + l.y + l.height)));

  it("anchors the footer's bottom edge at the footer distance from the page bottom", () => {
    const tree = layout(doc([para("body")], { footer: [para("F")], footerDistancePx: 30 }));
    expect(footerBottom(tree)).toBeCloseTo(SECTION.pageHeightPx - 30, 0);
  });

  it("centers the footer in the bottom margin when no distance is set", () => {
    const tree = layout(doc([para("body")], { footer: [para("F")] }));
    // sits inside the bottom margin, not pinned to a fixed distance.
    const bottom = footerBottom(tree);
    expect(bottom).toBeLessThan(SECTION.pageHeightPx);
    expect(bottom).toBeGreaterThan(SECTION.pageHeightPx - SECTION.marginPx.bottom);
  });

  it("reserves body space for an equation block in an anchored footer", () => {
    // Anchored footer (footerDistancePx) high enough that its top sits above the
    // bottom margin, so the band reservation drives contentBottomPx. A footer whose
    // only content is an equation must still count as visible (issue #13).
    const sec = { footerDistancePx: 200 };
    const eq = layout(doc([para("body")], { ...sec, footer: [equationBlock("<math><mfrac><mn>1</mn><mn>2</mn></mfrac></math>")] }));
    const blank = layout(doc([para("body")], { ...sec, footer: [para("")] }));
    // The equation reserves real height; a blank footer reserves nothing.
    expect(eq.pages[0]!.contentBottomPx).toBeLessThan(blank.pages[0]!.contentBottomPx);
  });
});

describe("engine — right indent", () => {
  it("narrows line wrapping by the right indent", () => {
    const long = "word ".repeat(40).trim();
    const wide = para(long);
    const narrow = para(long, { indentRightPx: 200 });
    const wideLines = placedOf(layout(doc([wide])), wide.id)!.pb.lines.length;
    const narrowLines = placedOf(layout(doc([narrow])), narrow.id)!.pb.lines.length;
    expect(narrowLines).toBeGreaterThan(wideLines);
  });

  it("pulls the right-aligned edge in by exactly the right indent", () => {
    const rightEdge = (p: Paragraph): number => {
      const line = placedOf(layout(doc([p])), p.id)!.pb.lines[0]!;
      return Math.max(...line.fragments.map((f) => f.x + f.width));
    };
    const plain = rightEdge(para("hi", { align: "right" }));
    const indented = rightEdge(para("hi", { align: "right", indentRightPx: 100 }));
    expect(plain - indented).toBeCloseTo(100, 0);
  });
});

// --- paragraph borders & shading (w:pBdr / w:shd) -------------------------

describe("engine — paragraph borders & shading", () => {
  const blue = { color: "#1a73e8", widthPx: 1 };

  it("attaches a paraDecor box spanning the content width minus indents", () => {
    const p = para("boxed paragraph", { borders: { top: blue, bottom: blue, left: blue, right: blue }, shading: "#eef4ff" });
    const pb = placedOf(layout(doc([p])), p.id)!.pb;
    expect(pb.paraDecor).toBeDefined();
    expect(pb.paraDecor!.shading).toBe("#eef4ff");
    expect(pb.paraDecor!.borders!.top!.color).toBe("#1a73e8");
    // content box is 816 - 96 - 96 = 624 wide, no indents.
    expect(pb.paraDecor!.width).toBeCloseTo(624, 0);
    // height equals the single line's height (lineHeight 1 → 16px char → 16px line).
    expect(pb.paraDecor!.height).toBeCloseTo(pb.lines[0]!.height, 5);
  });

  it("narrows the box by left + right indents", () => {
    const p = para("boxed", { borders: { left: blue }, indentLeftPx: 40, indentRightPx: 24 });
    const pb = placedOf(layout(doc([p])), p.id)!.pb;
    expect(pb.paraDecor!.width).toBeCloseTo(624 - 40 - 24, 0);
  });

  it("leaves paraDecor absent on a plain paragraph", () => {
    const p = para("plain");
    expect(placedOf(layout(doc([p])), p.id)!.pb.paraDecor).toBeUndefined();
  });

  it("reserves the top/bottom border widths in the block gap so the next block clears the box", () => {
    const top = para("first");
    const boxed = para("boxed", { borders: { top: { color: "#000", widthPx: 4 }, bottom: { color: "#000", widthPx: 4 } } });
    const after = para("after");
    const tree = layout(doc([top, boxed, after]));
    const boxedPb = placedOf(tree, boxed.id)!.pb;
    const afterPb = placedOf(tree, after.id)!.pb;
    // The follower starts at or below the box bottom plus the reserved bottom rule.
    expect(afterPb.y).toBeGreaterThanOrEqual(boxedPb.y + boxedPb.paraDecor!.height + 4 - 0.01);
  });

  it("decorates a bordered/shaded paragraph inside a table cell", () => {
    const boxed = para("cell box", { borders: { left: blue }, shading: "#eef4ff" });
    const t: TableBlock = {
      kind: "table", id: fresh(), revision: 0,
      rows: [{ cells: [{ id: fresh(), blocks: [boxed] }] }],
    };
    const tree = layout(doc([t]));
    const tablePb = placedOf(tree, t.id)!.pb;
    const cellPara = tablePb.table!.rows[0]!.cells[0]!.blocks.find((b) => b.blockId === boxed.id)!;
    expect(cellPara.paraDecor).toBeDefined();
    expect(cellPara.paraDecor!.shading).toBe("#eef4ff");
    expect(cellPara.paraDecor!.borders!.left!.color).toBe("#1a73e8");
    expect(cellPara.paraDecor!.width).toBeGreaterThan(0);
  });

  it("decorates a bordered/shaded paragraph in a header band", () => {
    const boxed = para("header box", { borders: { bottom: blue }, shading: "#eef4ff" });
    const tree = layout(doc([para("body")], { header: [boxed] }));
    const headerPb = tree.pages[0]!.header!.find((b) => b.blockId === boxed.id)!;
    expect(headerPb.paraDecor).toBeDefined();
    expect(headerPb.paraDecor!.shading).toBe("#eef4ff");
    expect(headerPb.paraDecor!.borders!.bottom!.color).toBe("#1a73e8");
  });
});

// --- floats (square-wrap images) ------------------------------------------

describe("engine — floats", () => {
  // A left-floated image: text flows beside it, y is not advanced.
  const leftFloat = (w: number, h: number): ImageBlock => ({ ...image(w, h), wrap: "square", align: "left" });

  it("justifies wrapped lines beside a float (not just left-aligns them)", () => {
    const img = leftFloat(200, 100);
    const body = para("word ".repeat(120).trim(), { align: "justify" });
    const pb = placedOf(layout(doc([img, body])), body.id)!.pb;
    // Lines whose fragments start well past the left margin are beside the float.
    const besideFloat = pb.lines.filter((l) => (l.fragments[0]?.x ?? 0) > 150);
    expect(besideFloat.length).toBeGreaterThan(1);
    // Justify stretches spaces via wordSpacingPx; a left fallback never sets it.
    const justifiedBesideFloat = besideFloat.some((l) => l.fragments.some((f) => (f.wordSpacingPx ?? 0) > 0));
    expect(justifiedBesideFloat).toBe(true);
  });

  it("leaves wrapped lines beside a float ragged when the paragraph is left-aligned", () => {
    const img = leftFloat(200, 100);
    const body = para("word ".repeat(120).trim(), { align: "left" });
    const pb = placedOf(layout(doc([img, body])), body.id)!.pb;
    const anyWordSpacing = pb.lines.some((l) => l.fragments.some((f) => (f.wordSpacingPx ?? 0) > 0));
    expect(anyWordSpacing).toBe(false);
  });

  it("drops a table below an active float instead of overlapping it", () => {
    const img = leftFloat(200, 100);
    const t = table([[cell("a"), cell("b")]]);
    const tree = layout(doc([img, t]));
    const imgPb = placedOf(tree, img.id)!.pb;
    const tablePb = placedOf(tree, t.id)!.pb;
    // The table must start at or below the float's bottom edge — never beside it.
    expect(tablePb.y).toBeGreaterThanOrEqual(imgPb.y + 100);
  });

  it("hangs a list marker beside a left float, not stranded on top of it", () => {
    const img = leftFloat(200, 100);
    const item = para("Item text that flows beside the floated image and wraps onward", {
      list: { listId: "numbers", level: 0 },
    });
    const d: Document = { ...doc([img, item]), lists: { numbers: defaultListDefinition("decimal") } };
    const tree = layout(d);
    const imgPb = placedOf(tree, img.id)!.pb;
    const imgRight = imgPb.x + imgPb.image!.width;
    const pb = placedOf(tree, item.id)!.pb;
    // The first line is pushed right past the float…
    expect(pb.lines[0]!.fragments[0]!.x).toBeGreaterThan(150);
    // …and the marker hangs with it, clear of the image — not back at the margin.
    expect(pb.marker).toBeDefined();
    expect(pb.marker!.x).toBeGreaterThanOrEqual(imgRight);
  });
});

// --- behind-text anchored images (wrapNone backgrounds) -------------------

describe("engine — behind-text anchored images", () => {
  // A full-page (816×1056) background anchored to the page origin, behind text —
  // the PARTY INVITATION.docx case. In Word it sits behind the text on one page.
  const bgImage = (): ImageBlock => ({
    ...image(816, 1056),
    anchor: { behind: true, offsetXPx: 0, offsetYPx: 0, relFromH: "page", relFromV: "page" },
  });

  it("keeps the whole document on one page (background consumes no flow height)", () => {
    const bg = bgImage();
    const body = ["Title", "Saturday, 21 August", "Address", "Please bring", "RSVP"].map((t) => para(t));
    const tree = layout(doc([bg, ...body]));
    expect(tree.pages).toHaveLength(1); // a flow-block 1056px image would force page 2
  });

  it("positions the background at its origin, behind text, first in the page block list", () => {
    const bg = bgImage();
    const title = para("Title");
    const tree = layout(doc([bg, title]));
    const page0 = tree.pages[0]!;
    expect(page0.blocks[0]!.blockId).toBe(bg.id); // first → painted under the text
    expect(page0.blocks[0]!.image!.behind).toBe(true);
    expect(page0.blocks[0]!.x).toBe(0); // page-relative origin + 0 offset
    expect(page0.blocks[0]!.y).toBe(0);
    // The title is NOT pushed down by the image's full-page height.
    const titlePb = placedOf(tree, title.id)!.pb;
    expect(titlePb.y).toBeLessThan(150); // near content top (96), not below a page of image
  });

  it("does not reflow surrounding text (registers no float)", () => {
    const bg = bgImage();
    const body = para("word ".repeat(120).trim());
    const withBg = placedOf(layout(doc([bg, body])), body.id)!.pb;
    const without = placedOf(layout(doc([body])), body.id)!.pb;
    // Same left edge and same wrapping as if the background weren't there.
    expect(withBg.lines[0]!.fragments[0]!.x).toBe(without.lines[0]!.fragments[0]!.x);
    expect(withBg.lines.length).toBe(without.lines.length);
  });

  it("honors a paragraph-relative negative offset (bleed toward the page corner)", () => {
    const bg: ImageBlock = {
      ...image(816, 1056),
      anchor: { behind: true, offsetXPx: -90, offsetYPx: -90, relFromH: "margin", relFromV: "paragraph" },
    };
    const tree = layout(doc([bg, para("Title")]));
    const pb = tree.pages[0]!.blocks[0]!;
    // margin origin (left/top = 96) shifted by -90 → 6, just inside the page corner.
    expect(pb.x).toBe(6);
    expect(pb.y).toBe(6);
  });

  // A front-of-text anchor authored BEFORE the text must still paint over it.
  const frontImage = (z = 0): ImageBlock => ({
    ...image(200, 100),
    anchor: { behind: false, offsetXPx: 0, offsetYPx: 0, relFromH: "page", relFromV: "page", z },
  });

  it("moves an in-front anchored image to the END of the page block list (paints over text)", () => {
    const fg = frontImage();
    const body = para("hello");
    const tree = layout(doc([fg, body])); // authored first…
    const blocks = tree.pages[0]!.blocks;
    expect(blocks[blocks.length - 1]!.blockId).toBe(fg.id); // …but painted last (on top)
    expect(blocks[blocks.length - 1]!.image!.front).toBe(true);
  });

  it("orders the layers behind → flow → front regardless of document order", () => {
    const behind = bgImage();
    const front = frontImage();
    const body = para("middle");
    // Document order deliberately scrambled: front, body, behind.
    const tree = layout(doc([front, body, behind]));
    const ids = tree.pages[0]!.blocks.map((b) => b.blockId);
    expect(ids.indexOf(behind.id)).toBeLessThan(ids.indexOf(body.id));
    expect(ids.indexOf(body.id)).toBeLessThan(ids.indexOf(front.id));
  });

  it("stacks same-layer anchored images by z (higher paints later)", () => {
    const lo = frontImage(1);
    const hi = frontImage(5);
    // hi authored first, but its higher z must place it after lo.
    const tree = layout(doc([hi, lo, para("x")]));
    const ids = tree.pages[0]!.blocks.map((b) => b.blockId);
    expect(ids.indexOf(lo.id)).toBeLessThan(ids.indexOf(hi.id));
  });

  it("leaves the block order untouched when no anchored images exist (no PDF drift)", () => {
    const a = para("a");
    const t = table([[cell("x")]]);
    const b = para("b");
    const tree = layout(doc([a, t, b]));
    // Plain flow documents must keep exact document order — the z-order pass is
    // skipped entirely, so canvas and PDF output are byte-identical to before.
    expect(tree.pages[0]!.blocks.map((bl) => bl.blockId)).toEqual([a.id, t.id, b.id]);
  });

  // --- anchored images INSIDE table cells ---------------------------------

  const cellBehindImage = (): ImageBlock => ({
    ...image(120, 120),
    anchor: { behind: true, offsetXPx: 5, offsetYPx: 7, relFromH: "column", relFromV: "paragraph" },
  });

  it("carries the behind flag and lifts a cell anchored image into the cell's behind layer", () => {
    const bg = cellBehindImage();
    const txt = para("over the image");
    const t: TableBlock = { ...table([[cell("placeholder")]]) };
    t.rows[0]!.cells[0]!.blocks = [bg, txt];
    const tree = layout(doc([t]));
    const cellBlocks = tree.pages[0]!.blocks[0]!.table!.rows[0]!.cells[0]!.blocks;
    const imgIdx = cellBlocks.findIndex((b) => b.image?.behind);
    const textIdx = cellBlocks.findIndex((b) => b.lines.length > 0);
    expect(imgIdx).toBeGreaterThanOrEqual(0);
    expect(cellBlocks[imgIdx]!.image!.behind).toBe(true);
    expect(cellBlocks[imgIdx]!.image!.width).toBe(120); // native size, not scaled to cell
    expect(imgIdx).toBeLessThan(textIdx); // behind layer paints before the text
  });

  it("does not consume cell height for a behind cell image (text sits at the cell top)", () => {
    const textY = (blocks: Block[]): number => {
      const t: TableBlock = { ...table([[cell("placeholder")]]) };
      t.rows[0]!.cells[0]!.blocks = blocks;
      const cellBlocks = layout(doc([t])).pages[0]!.blocks[0]!.table!.rows[0]!.cells[0]!.blocks;
      return cellBlocks.find((b) => b.lines.length > 0)!.y;
    };
    // The 120px behind image must add no vertical space: the text lands at the
    // same y whether or not the behind image precedes it.
    expect(textY([cellBehindImage(), para("hello")])).toBe(textY([para("hello")]));
  });

  it("leaves a cell's block order untouched when it has no anchored images (no PDF drift)", () => {
    // An ordinary image cell must keep document order and scale to the cell — the
    // recursive z-order pass is a no-op without behind/front images in the cell.
    const flow = image(2000, 1000); // wider than the cell → scaled down in-flow
    const txt = para("caption");
    const t: TableBlock = { ...table([[cell("placeholder")]]) };
    t.rows[0]!.cells[0]!.blocks = [flow, txt];
    const cellBlocks = layout(doc([t])).pages[0]!.blocks[0]!.table!.rows[0]!.cells[0]!.blocks;
    expect(cellBlocks[0]!.blockId).toBe(flow.id); // unchanged order
    expect(cellBlocks[0]!.image!.behind).toBeUndefined();
    expect(cellBlocks[0]!.image!.width).toBeLessThan(2000); // still scaled to fit the cell
    expect(cellBlocks[1]!.blockId).toBe(txt.id);
  });

  it("makes a behind image selectable only where no text covers it", () => {
    const bg = bgImage(); // full page (0,0)–(816,1056), behind text
    const title = para("hello");
    const tree = layout(doc([bg, title]));
    const titlePb = placedOf(tree, title.id)!.pb;

    // Foreground hit-testing never returns the background.
    expect(hitTestObject(tree, 0, 400, 400)).toBeNull();
    // A click on the text falls through (caret), selecting no object.
    expect(hitTestSelectableObject(tree, 0, titlePb.x + 4, titlePb.y + 2)).toBeNull();
    // A click on the bare background (no text under it) selects the background.
    expect(hitTestSelectableObject(tree, 0, 400, titlePb.y + 300)?.blockId).toBe(bg.id);
    // A background bleeding into the TOP MARGIN (above the content box) is still
    // found there — the controller must not let the header band gate it out.
    expect(30).toBeLessThan(tree.pages[0]!.contentTopPx);
    expect(hitTestSelectableObject(tree, 0, 400, 30)?.blockId).toBe(bg.id);
  });
});

// --- formatting marks (show/hide non-printing marks) -----------------------
// The overlay paint reads these per-line flags from the engine plus spaceMarkXs
// from geometry; both are pure and font-stub-deterministic, so they pin the
// feature without exercising the canvas renderer's DOM.

describe("engine — formatting marks", () => {
  const linesOf = (p: Paragraph): import("./layoutTree").LineBox[] =>
    placedOf(layout(doc([p])), p.id)!.pb.lines;

  it("flags the paragraph's final line with paragraphEnd (single line, no break)", () => {
    const lines = linesOf(para("Hello world"));
    expect(lines.length).toBe(1);
    expect(lines[0]!.paragraphEnd).toBe(true);
    expect(lines[0]!.lineBreak).toBeFalsy();
  });

  it("flags an empty paragraph's sole line with paragraphEnd (Word shows a bare ¶)", () => {
    const lines = linesOf(para(""));
    expect(lines[lines.length - 1]!.paragraphEnd).toBe(true);
  });

  it("sets paragraphEnd ONLY on the last line of a multi-line paragraph", () => {
    const lines = linesOf(para("word ".repeat(60).trim()));
    expect(lines.length).toBeGreaterThan(1);
    expect(lines[lines.length - 1]!.paragraphEnd).toBe(true);
    expect(lines.slice(0, -1).every((l) => !l.paragraphEnd)).toBe(true);
    // Wrapped lines are NOT manual breaks — no stray ↵ on soft-wrapped lines.
    expect(lines.every((l) => !l.lineBreak)).toBe(true);
  });

  it("marks a soft line break (\\v) with lineBreak and the paragraph end with paragraphEnd", () => {
    const lines = linesOf(para("Line one\vLine two"));
    const breaks = lines.filter((l) => l.lineBreak);
    expect(breaks.length).toBe(1); // exactly the segment boundary
    expect(breaks[0]!.paragraphEnd).toBeFalsy(); // a break is not a paragraph end
    const last = lines[lines.length - 1]!;
    expect(last.paragraphEnd).toBe(true);
    expect(last.lineBreak).toBeFalsy();
  });

  it("emits a tab arrow in the gap a tab opened, ending at the positioned piece", () => {
    const p = para("Left\tRight", { tabStops: [{ posPx: 300, align: "right", leader: "dot" }] });
    const line = placedOf(layout(doc([p])), p.id)!.pb.lines[0]!;
    expect(line.tabArrows?.length).toBe(1);
    const arrow = line.tabArrows![0]!;
    expect(arrow.x2).toBeGreaterThan(arrow.x1); // a real gap
    // The arrow gap closes exactly where the right-aligned piece begins.
    const right = line.fragments.find((f) => f.text === "Right")!;
    expect(arrow.x2).toBeCloseTo(right.x, 0);
  });

  it("emits a tab arrow even for a plain (leaderless) tab", () => {
    const p = para("A\tB");
    const line = placedOf(layout(doc([p])), p.id)!.pb.lines[0]!;
    expect(line.leaders).toBeUndefined(); // default tab stop has no leader…
    expect(line.tabArrows?.length).toBe(1); // …but the arrow is independent of leaders
  });

  it("spaceMarkXs returns one mid-glyph center per rendered space, left to right", () => {
    const p = para("a b c"); // two interior spaces at indices 1 and 3
    const frag = placedOf(layout(doc([p])), p.id)!.pb.lines[0]!.fragments[0]!;
    const xs = spaceMarkXs(frag);
    expect(xs.length).toBe(2);
    expect(xs[0]!).toBeLessThan(xs[1]!);
    // Stub font: 16px → 8px/char. Space centers sit at frag.x + index*8 + 4.
    expect(xs[0]!).toBeCloseTo(frag.x + 1 * 8 + 4, 0);
    expect(xs[1]!).toBeCloseTo(frag.x + 3 * 8 + 4, 0);
  });
});

describe("engine — autofit tables", () => {
  // Stub font: 16px → 8px/char; default cell margin adds 7.2px L+R = 14.4 pad.
  // Content box is 624px wide (816 − 96 − 96).
  const sum = (xs: number[]): number => xs.reduce((s, x) => s + x, 0);
  const autofit = (rows: TableCell[][], mode: "autofitContents" | "autofitWindow"): TableBlock => ({
    ...table(rows),
    widthMode: mode,
  });

  it("AutoFit to Contents shrinks the table below the page width and sizes columns by content", () => {
    const t = autofit(
      [
        [cell("ID"), cell("Description")],
        [cell("1"), cell("short")],
        [cell("42"), cell("a much longer description cell")],
      ],
      "autofitContents",
    );
    const placed = placedOf(layout(doc([t])), t.id)!.pb.table!;
    // ID col ≈ "ID"(16) + 14.4 = 30.4; the table is far narrower than the 624 box.
    expect(placed.colWidths[0]!).toBeCloseTo(30.4, 0);
    expect(placed.colWidths[0]!).toBeLessThan(placed.colWidths[1]!);
    expect(sum(placed.colWidths)).toBeLessThan(400);
    expect(placed.width).toBeLessThan(400); // placed table width follows the solved total
  });

  it("AutoFit to Window solves by content then fills the available width", () => {
    const t = autofit(
      [
        [cell("ID"), cell("Description")],
        [cell("42"), cell("a much longer description cell")],
      ],
      "autofitWindow",
    );
    const placed = placedOf(layout(doc([t])), t.id)!.pb.table!;
    expect(sum(placed.colWidths)).toBeCloseTo(624, 0); // exactly fills the content box
    expect(placed.colWidths[1]!).toBeGreaterThan(placed.colWidths[0]!); // proportions preserved
  });

  it("fixed tables (no widthMode) keep the exact colFractions × content-width layout", () => {
    const t = table([[cell("a"), cell("b"), cell("c")]]); // equal thirds, no widthMode
    const placed = placedOf(layout(doc([t])), t.id)!.pb.table!;
    for (const w of placed.colWidths) expect(w).toBeCloseTo(208, 0); // 624 / 3
    expect(sum(placed.colWidths)).toBeCloseTo(624, 0);
  });

  it("fixed table with a px preferred width sizes the whole grid to that width", () => {
    const t: TableBlock = { ...table([[cell("a"), cell("b")]]), preferredWidth: { type: "px", value: 300 } };
    const placed = placedOf(layout(doc([t])), t.id)!.pb.table!;
    expect(placed.width).toBeCloseTo(300, 0);
    expect(sum(placed.colWidths)).toBeCloseTo(300, 0);
    for (const w of placed.colWidths) expect(w).toBeCloseTo(150, 0); // proportions preserved
  });

  it("fixed table with a percent preferred width sizes to that fraction of the content box", () => {
    const t: TableBlock = { ...table([[cell("a"), cell("b")]]), preferredWidth: { type: "pct", value: 50 } };
    const placed = placedOf(layout(doc([t])), t.id)!.pb.table!;
    expect(placed.width).toBeCloseTo(312, 0); // 624 * 0.5
  });

  it("clamps a preferred width wider than the content box back to the content box", () => {
    const t: TableBlock = { ...table([[cell("a"), cell("b")]]), preferredWidth: { type: "px", value: 5000 } };
    const placed = placedOf(layout(doc([t])), t.id)!.pb.table!;
    expect(placed.width).toBeCloseTo(624, 0);
  });

  it("center-aligns a narrow table by offsetting its x within the content band", () => {
    const t: TableBlock = {
      ...table([[cell("a"), cell("b")]]),
      preferredWidth: { type: "px", value: 300 },
      align: "center",
    };
    const pb = placedOf(layout(doc([t])), t.id)!.pb;
    expect(pb.table!.x).toBeCloseTo(96 + (624 - 300) / 2, 0); // left margin + half the slack
    expect(pb.x).toBeCloseTo(96 + (624 - 300) / 2, 0);
  });

  it("right-aligns a narrow table flush to the content band's right edge", () => {
    const t: TableBlock = {
      ...table([[cell("a"), cell("b")]]),
      preferredWidth: { type: "px", value: 300 },
      align: "right",
    };
    const placed = placedOf(layout(doc([t])), t.id)!.pb.table!;
    expect(placed.x).toBeCloseTo(96 + (624 - 300), 0);
    expect(placed.x + placed.width).toBeCloseTo(96 + 624, 0);
  });

  it("aligns an AutoFit-to-Contents table too (alignment is independent of width mode)", () => {
    const t: TableBlock = { ...autofit([[cell("ID"), cell("x")]], "autofitContents"), align: "right" };
    const placed = placedOf(layout(doc([t])), t.id)!.pb.table!;
    expect(placed.x).toBeGreaterThan(96); // shifted right of the left margin
    expect(placed.x + placed.width).toBeCloseTo(96 + 624, 0); // flush right edge
  });

  it("distributes a colSpan cell's demand across the columns it covers without breaking the grid", () => {
    const t = autofit(
      [
        [cell("A wide spanning header", { colSpan: 3 })],
        [cell("a"), cell("bb"), cell("ccc")],
      ],
      "autofitContents",
    );
    const placed = placedOf(layout(doc([t])), t.id)!.pb.table!;
    expect(placed.colWidths).toHaveLength(3);
    expect(placed.colWidths.every((w) => Number.isFinite(w) && w > 0)).toBe(true);
  });

  it("clamps a column up to a cell's preferred width (w:tcW)", () => {
    const t = autofit(
      [[cell("x", { preferredWidth: { px: 300, type: "abs" } }), cell("y")]],
      "autofitContents",
    );
    const placed = placedOf(layout(doc([t])), t.id)!.pb.table!;
    expect(placed.colWidths[0]!).toBeGreaterThanOrEqual(299); // preference wins over the tiny "x"
  });
});

describe("engine — table measure cache", () => {
  const long = "word ".repeat(60); // wraps to several lines in a content-width cell

  it("re-measures when a cell paragraph grows (revision bump invalidates)", () => {
    const eng = createLayoutEngine();
    const t = table([[cell("short")]], [1]);
    const h1 = placedOf(eng.layout(doc([t])), t.id)!.pb.table!.height;
    // Edit the cell paragraph exactly as replaceParagraphAt would: longer runs +
    // bumped paragraph AND table revisions.
    const cp = t.rows[0]!.cells[0]!.blocks[0] as Paragraph;
    const t2: TableBlock = {
      ...t,
      revision: t.revision + 1,
      rows: [{ cells: [{ ...t.rows[0]!.cells[0]!, blocks: [{ ...cp, revision: cp.revision + 1, runs: [{ text: long, style: CHAR }] }] }] }],
    };
    const h2 = placedOf(eng.layout(doc([t2])), t.id)!.pb.table!.height;
    expect(h2).toBeGreaterThan(h1); // not stale — the table re-measured taller
  });

  it("produces identical geometry across re-layouts of an unchanged table", () => {
    const eng = createLayoutEngine();
    const t = table([[cell("a"), cell("b")]], [0.5, 0.5]);
    const d = doc([para("intro"), t]);
    const a = placedOf(eng.layout(d), t.id)!.pb.table!;
    const b = placedOf(eng.layout(d), t.id)!.pb.table!;
    expect(b.height).toBe(a.height);
    expect(b.colWidths).toEqual(a.colWidths);
  });

  it("invalidates on a column-width (section width) change", () => {
    const eng = createLayoutEngine();
    const t = table([[cell("a"), cell("b")]], [0.5, 0.5]);
    const wide = placedOf(eng.layout(doc([t])), t.id)!.pb.table!;
    const narrow = placedOf(eng.layout(doc([t], { pageWidthPx: 560 })), t.id)!.pb.table!;
    expect(narrow.colWidths[0]!).toBeLessThan(wide.colWidths[0]!);
  });

  it("prunes a removed table so a re-minted same-id table re-measures", () => {
    const eng = createLayoutEngine();
    const t = table([[cell(long)]], [1]);
    const p = para("keep");
    const tall = placedOf(eng.layout(doc([p, t])), t.id)!.pb.table!.height;
    eng.layout(doc([p])); // t absent this pass → pruned from the table cache
    // Same id at revision 0 but short content; without the prune the cache would
    // serve the stale tall measurement (same id+revision+width).
    const reminted: TableBlock = { ...t, rows: [{ cells: [{ ...t.rows[0]!.cells[0]!, blocks: [para("hi")] }] }] };
    const short = placedOf(eng.layout(doc([p, reminted])), t.id)!.pb.table!.height;
    expect(short).toBeLessThan(tall);
  });

  it("re-resolves a cached table's PAGE field when the table moves to a new page (#48)", () => {
    const eng = createLayoutEngine();
    const fieldDefs: Record<string, FieldDef> = {
      pg: { id: "pg", instruction: " PAGE ", name: "PAGE", kind: "builtin", spec: { type: "PAGE" } },
    };
    // A table whose only cell holds a live PAGE field — the run text is the {page}
    // token, tagged with the builtin PAGE field def the engine resolves per page.
    const fieldPara: Paragraph = {
      kind: "paragraph", id: fresh(), revision: 0,
      runs: [{ text: "{page}", style: { ...CHAR, fieldId: "pg" } }], style: PARA,
    };
    const t = table([[{ id: fresh(), blocks: [fieldPara] }]], [1]);
    const fieldText = (tree: ReturnType<typeof layout>): string =>
      placedOf(tree, t.id)!.pb.table!.rows[0]!.cells[0]!.blocks[0]!.lines
        .flatMap((l) => l.fragments.map((f) => f.text)).join("");

    // Pass 1: the table sits on page 1 → field resolves to "1".
    const t1 = eng.layout({ ...doc([t]), fields: fieldDefs });
    expect(placedOf(t1, t.id)!.page).toBe(0);
    expect(fieldText(t1)).toBe("1");

    // Pass 2: a page break before the SAME (unchanged) table pushes it to page 2.
    // Its revision+width are unchanged → tableCache hit. Clone-on-write keeps the
    // cached {page} token pristine, so the field re-resolves to "2" instead of
    // serving the baked "1" the in-place mutation used to leave behind.
    const t2 = eng.layout({ ...doc([para("intro"), para("brk", { pageBreakBefore: true }), t]), fields: fieldDefs });
    expect(placedOf(t2, t.id)!.page).toBe(1);
    expect(fieldText(t2)).toBe("2");

    // Pass 3: back to page 1 — the cache must not have been corrupted either way.
    const t3 = eng.layout({ ...doc([t]), fields: fieldDefs });
    expect(fieldText(t3)).toBe("1");
  });
});

describe("engine — band (header/footer) layout cache", () => {
  const footerText = (tree: ReturnType<typeof layout>, pageIndex: number): string =>
    (tree.pages[pageIndex]!.footer ?? []).flatMap((b) => b.lines.flatMap((l) => l.fragments.map((f) => f.text))).join("");
  const tall = (): Block[] => Array.from({ length: 160 }, (_, i) => para(`body line ${i}`)); // spans several pages

  it("substitutes the correct page number per page (cache keyed by pageNum)", () => {
    const eng = createLayoutEngine();
    const d = doc(tall(), { footer: [para("{page}")] });
    const t = eng.layout(d);
    expect(t.pages.length).toBeGreaterThan(2);
    t.pages.forEach((_, i) => expect(footerText(t, i)).toContain(String(i + 1)));
    // Re-layout hits the band cache and must still be correct per page.
    const t2 = eng.layout(d);
    t2.pages.forEach((_, i) => expect(footerText(t2, i)).toContain(String(i + 1)));
  });

  it("a body edit reuses the (unchanged) band cache without corrupting page numbers", () => {
    const eng = createLayoutEngine();
    const footer = [para("Page {page}")];
    const body = tall();
    const t1 = eng.layout(doc(body, { footer }));
    // Edit a body paragraph (bump revision); the SAME footer array → band cache hit.
    const body2 = body.slice();
    const bp = body2[0] as Paragraph;
    body2[0] = { ...bp, revision: bp.revision + 1, runs: [{ text: "edited", style: CHAR }] };
    const t2 = eng.layout(doc(body2, { footer }));
    expect(footerText(t2, 1)).toBe(footerText(t1, 1)); // page 2 footer unchanged + correct
    expect(footerText(t2, 1)).toContain("2");
  });

  it("invalidates when the footer blocks change (new array identity)", () => {
    const eng = createLayoutEngine();
    const body = tall();
    const a = footerText(eng.layout(doc(body, { footer: [para("{page}")] })), 0);
    const b = footerText(eng.layout(doc(body, { footer: [para("p{page}")] })), 0);
    expect(a).toBe("1");
    expect(b).toBe("p1");
  });
});
