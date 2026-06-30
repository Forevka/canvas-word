// All mutation flows through operations. Applying an op returns the new document
// (structurally shared), the op's exact INVERSE (undo is free), and a position
// mapper so selections survive edits. This seam is also where OT/CRDT
// collaboration would slot in later.

import type {
  BandContainer,
  Block,
  BookmarkRange,
  CharStyle,
  Document,
  EquationBlock,
  FieldDef,
  ImageBlock,
  ParaStyle,
  Paragraph,
  Run,
  RowProps,
  SdtProps,
  TableBlock,
  TableCell,
  TableRow,
} from "./document";
import { BAND_CONTAINERS } from "./document";
import type { MathEquation } from "./math";
import { sdtPathEq } from "./sdt";
import { DEFAULT_CHAR_STYLE } from "./defaults";
import type { DocPosition } from "./position";
import {
  locateParagraph,
  paragraphAt,
  replaceParagraphAt,
  styleAtRuns,
  textOfRuns,
  type ParaLocation,
} from "./text";

/** setImageProps payload. `wrap` and `anchor` are mutually-exclusive states, so
 *  each accepts `null` as an explicit "clear this field" sentinel (plain
 *  `undefined` means "leave unchanged", which exactOptionalPropertyTypes also
 *  needs to keep these distinguishable). */
export interface ImagePropsPatch {
  widthPx?: number;
  heightPx?: number;
  align?: ImageBlock["align"];
  wrap?: ImageBlock["wrap"] | null;
  anchor?: ImageBlock["anchor"] | null;
  /** Crop insets (a:srcRect 0..1 fractions). `null` clears the crop. */
  crop?: ImageBlock["crop"] | null;
}

/** setTableProps payload — table-LEVEL fields (w:tblPr): indent + the cascade
 *  defaults (borders/shading/cell margins). Each key follows the setImageProps
 *  convention: a value sets the field, `null` clears it, and an absent key leaves
 *  it unchanged (exactOptionalPropertyTypes keeps the two distinguishable). */
export interface TablePropsPatch {
  indentPx?: number | null;
  defaultBorders?: import("./document").TableBorders | null;
  defaultShading?: string | null;
  defaultCellMargin?: import("./document").CellMargin | null;
}

/** Clamp crop insets to the documented invariant — each in [0,1) with a non-empty
 *  visible window — so a malformed local/remote patch can't persist impossible
 *  values (negative, NaN, or a fully-collapsed window) in the shared model. */
function normalizeCrop(c: NonNullable<ImageBlock["crop"]>): NonNullable<ImageBlock["crop"]> {
  const clamp = (v: number): number => (Number.isFinite(v) ? Math.min(0.999, Math.max(0, v)) : 0);
  let { left, top, right, bottom } = { left: clamp(c.left), top: clamp(c.top), right: clamp(c.right), bottom: clamp(c.bottom) };
  if (left + right >= 1) {
    const k = 0.999 / (left + right);
    left *= k;
    right *= k;
  }
  if (top + bottom >= 1) {
    const k = 0.999 / (top + bottom);
    top *= k;
    bottom *= k;
  }
  return { left, top, right, bottom };
}

export type Op =
  | { type: "insertText"; at: DocPosition; text: string; style?: CharStyle }
  | { type: "insertRuns"; at: DocPosition; runs: Run[] }
  | { type: "deleteRange"; blockId: string; start: number; end: number }
  | { type: "setRuns"; blockId: string; runs: Run[] }
  | { type: "setParaStyle"; blockId: string; patch: Partial<ParaStyle> }
  | { type: "splitParagraph"; at: DocPosition; newBlockId: string; newStyle?: ParaStyle; newSdtPath?: string[] }
  | { type: "mergeParagraphs"; firstBlockId: string }
  | { type: "insertBlock"; index: number; block: Block; where?: Container }
  | { type: "removeBlock"; blockId: string }
  | { type: "setImageProps"; blockId: string; patch: ImagePropsPatch }
  | { type: "setEquation"; blockId: string; equation: MathEquation }
  | { type: "setEquationAlign"; blockId: string; align: "left" | "center" | "right" }
  | { type: "setTableRow"; tableId: string; rowIndex: number; row: TableRow }
  | { type: "setTableStructure"; tableId: string; rows: TableRow[]; colFractions?: number[] }
  | { type: "setTableStyleRef"; tableId: string; styleId: string | null; condOverrides?: import("./document").TableCondOverrides | null }
  | { type: "setTableColFractions"; blockId: string; fractions: number[] }
  | { type: "setTableWidthMode"; blockId: string; mode: TableBlock["widthMode"] }
  | { type: "setTablePreferredWidth"; blockId: string; width: TableBlock["preferredWidth"] | null }
  | { type: "setTableAlign"; blockId: string; align: TableBlock["align"] | null }
  | { type: "setTableProps"; blockId: string; patch: TablePropsPatch }
  | { type: "insertTableRow"; tableId: string; rowIndex: number; row: TableRow }
  | { type: "removeTableRow"; tableId: string; rowIndex: number }
  | { type: "setRowHeight"; tableId: string; rowIndex: number; height: NonNullable<RowProps["height"]> | null }
  | { type: "insertTableColumn"; tableId: string; colIndex: number; cells: TableCell[]; fractions?: number[] }
  | { type: "removeTableColumn"; tableId: string; colIndex: number }
  | { type: "setStylesheet"; stylesheet: import("./stylesheet").Stylesheet }
  | { type: "setTableStyleSheet"; tableStyles: Record<string, import("./tableStyles").TableStyle> }
  | { type: "setListDefinition"; listId: string; def: import("./lists").ListDefinition | null }
  | { type: "setSectionProps"; geometry: SectionGeometry }
  | { type: "setSectionBand"; band: BandContainer; blocks: Block[] | null }
  | { type: "setFootnote"; noteId: string; paras: Paragraph[] | null }
  | { type: "setEndnote"; noteId: string; paras: Paragraph[] | null }
  | { type: "setSdtProps"; id: string; props: SdtProps | null }
  | { type: "setField"; id: string; def: FieldDef | null }
  | { type: "setTocInstruction"; instruction: string | null }
  | { type: "setBookmark"; name: string; range: BookmarkRange | null };

/** Page-setup fields of the final section (`doc.section`). Bands are NOT here —
 *  they change through container ops; mid-document sections change through
 *  setParaStyle on their break paragraph's `sectionBreak.props`. */
export interface SectionGeometry {
  pageWidthPx: number;
  pageHeightPx: number;
  marginPx: { top: number; right: number; bottom: number; left: number };
  /** `null` = single column (the explicit "off" — SectionPatch distinguishes it
   *  from "inherit"). `sep`/`cols` carry the separator line and per-column
   *  widths respectively. */
  columns: { count: number; gapPx: number; sep?: boolean; cols?: import("./document").ColumnEntry[] } | null;
  /** `null` = continue numbering from the previous section. */
  pageNumberStart: number | null;
  /** `null` = inherit (center band in margin). px from page top to header top. */
  headerDistancePx: number | null;
  /** `null` = inherit (center band in margin). px from page bottom to footer bottom. */
  footerDistancePx: number | null;
  /** `null` = no page fill. "#rrggbb". */
  pageColorHex: string | null;
  /** `null` = no page border. */
  pageBorders: import("./document").PageBorders | null;
  /** OOXML w:sectPr/w:type for the section START. `null` = the default ("nextPage");
   *  "evenPage"/"oddPage" force the section's first page onto an even/odd page. */
  breakType: import("./document").SectionBreakType | null;
  /** w:sectPr/w:lnNumType — line numbering in the margin. `null` = off. */
  lineNumbering: import("./document").LineNumbering | null;
}

/** Top-level block containers: the body, or one of the six margin-band stories
 *  (default header/footer + first/even variants). */
export type Container = "body" | BandContainer;

export interface ApplyResult {
  doc: Document;
  inverse: Op;
  /** Remaps any stored position across this edit (selection, bookmarks). */
  mapPosition(pos: DocPosition): DocPosition;
  dirtyBlockIds: string[];
}

// ---------------------------------------------------------------------------
// Run-list surgery (pure helpers)

/** Compare two w:sym markers by value (font + code point). */
function symbolEq(a: CharStyle["symbol"], b: CharStyle["symbol"]): boolean {
  if (!a || !b) return !a === !b;
  return a.font === b.font && a.char === b.char;
}

export function styleEq(a: CharStyle, b: CharStyle): boolean {
  return (
    a.fontFamily === b.fontFamily &&
    a.fontFamilyComplexScript === b.fontFamilyComplexScript &&
    a.fontFamilyEastAsia === b.fontFamilyEastAsia &&
    a.fontSizePx === b.fontSizePx &&
    a.bold === b.bold &&
    a.italic === b.italic &&
    a.underline === b.underline &&
    a.underlineStyle === b.underlineStyle &&
    a.underlineColor === b.underlineColor &&
    a.strikethrough === b.strikethrough &&
    a.color === b.color &&
    !!a.hidden === !!b.hidden && // hidden runs must never merge with visible ones
    (a.letterSpacingPx ?? 0) === (b.letterSpacingPx ?? 0) &&
    a.highlightColor === b.highlightColor &&
    a.verticalAlign === b.verticalAlign &&
    !!a.caps === !!b.caps && // case transforms change the rendered glyphs — never merge across
    !!a.smallCaps === !!b.smallCaps &&
    a.link === b.link &&
    a.footnoteRef === b.footnoteRef && // adjacent refs must never merge into one run
    sdtPathEq(a.sdtPath, b.sdtPath) && // content-control boundaries (incl. nesting) survive normalization
    a.fieldId === b.fieldId && // inline-field boundaries survive normalization
    a.equation === b.equation && // inline equations are atomic — never merge (reference identity)
    // Minor run typography & effects (w:rPr extras) — compared so styled runs never
    // merge with plain ones (and a struck-through pair keeps its boundary).
    !!a.doubleStrikethrough === !!b.doubleStrikethrough &&
    (a.positionPx ?? 0) === (b.positionPx ?? 0) &&
    (a.kerningMinPx ?? 0) === (b.kerningMinPx ?? 0) &&
    (a.widthScalePct ?? 100) === (b.widthScalePct ?? 100) &&
    a.emphasisMark === b.emphasisMark &&
    !!a.outline === !!b.outline &&
    !!a.shadow === !!b.shadow &&
    !!a.emboss === !!b.emboss &&
    !!a.imprint === !!b.imprint &&
    (a.fitTextPx ?? 0) === (b.fitTextPx ?? 0) &&
    runBorderEq(a.runBorder, b.runBorder) &&
    symbolEq(a.symbol, b.symbol) // symbol glyphs (w:sym) carry font+codepoint — never merge with text
  );
}

/** Structural equality for a run border (w:bdr) — runs with differently-styled
 *  borders must not merge, but two imported runs with identical borders may. */
function runBorderEq(a: CharStyle["runBorder"], b: CharStyle["runBorder"]): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  return a.color === b.color && a.widthPx === b.widthPx && a.style === b.style;
}

/** Merge equal-styled neighbors, drop empties. An all-empty paragraph keeps ONE
 *  empty run so the paragraph mark still carries a style (Word behavior). */
export function normalizeRuns(runs: Run[], fallback: CharStyle): Run[] {
  const nonEmpty = runs.filter((r) => r.text.length > 0);
  if (nonEmpty.length === 0) return [{ text: "", style: runs[0]?.style ?? fallback }];
  const out: Run[] = [{ ...nonEmpty[0]! }];
  for (let i = 1; i < nonEmpty.length; i++) {
    const r = nonEmpty[i]!;
    const last = out[out.length - 1]!;
    if (styleEq(last.style, r.style)) last.text += r.text;
    else out.push({ ...r });
  }
  return out;
}

/** Split a run list at a UTF-16 offset. */
export function splitRunsAt(runs: Run[], offset: number): [Run[], Run[]] {
  const head: Run[] = [];
  const tail: Run[] = [];
  let cum = 0;
  for (const r of runs) {
    const end = cum + r.text.length;
    if (end <= offset) head.push(r);
    else if (cum >= offset) tail.push(r);
    else {
      head.push({ text: r.text.slice(0, offset - cum), style: r.style });
      tail.push({ text: r.text.slice(offset - cum), style: r.style });
    }
    cum = end;
  }
  return [head, tail];
}

export function sliceRuns(runs: Run[], start: number, end: number): Run[] {
  const [, fromStart] = splitRunsAt(runs, start);
  const [middle] = splitRunsAt(fromStart, end - start);
  return middle;
}

function fallbackStyle(runs: Run[], offset: number): CharStyle {
  return styleAtRuns(runs, offset) ?? { ...DEFAULT_CHAR_STYLE };
}

export function insertTextInRuns(runs: Run[], offset: number, text: string, style?: CharStyle): Run[] {
  const st = style ?? fallbackStyle(runs, offset);
  const [head, tail] = splitRunsAt(runs, offset);
  return normalizeRuns([...head, { text, style: st }, ...tail], st);
}

export function insertRunsInRuns(runs: Run[], offset: number, inserted: Run[]): Run[] {
  const [head, tail] = splitRunsAt(runs, offset);
  return normalizeRuns([...head, ...inserted, ...tail], fallbackStyle(runs, offset));
}

export function deleteInRuns(runs: Run[], start: number, end: number): Run[] {
  const [head, fromStart] = splitRunsAt(runs, start);
  const [, tail] = splitRunsAt(fromStart, end - start);
  return normalizeRuns([...head, ...tail], fallbackStyle(runs, start));
}

export function applyStylePatchToRuns(
  runs: Run[],
  start: number,
  end: number,
  patch: Partial<CharStyle>,
): Run[] {
  const [head, fromStart] = splitRunsAt(runs, start);
  const [middle, tail] = splitRunsAt(fromStart, end - start);
  const styled = middle.map((r) => ({ text: r.text, style: { ...r.style, ...patch } }));
  return normalizeRuns([...head, ...styled, ...tail], fallbackStyle(runs, start));
}

/** Rewrite the TEXT of runs in [start,end) through `fn`, preserving each run's
 *  style. `fn` receives the whole covered string (so it can do context-aware
 *  transforms like sentence case) and must return a string of the SAME length —
 *  if the length changes the slice is left untouched (run offsets must stay
 *  stable). Used by change-case; transforms here are 1:1 per character. */
export function mapTextInRuns(
  runs: Run[],
  start: number,
  end: number,
  fn: (covered: string) => string,
): Run[] {
  const [head, fromStart] = splitRunsAt(runs, start);
  const [middle, tail] = splitRunsAt(fromStart, end - start);
  const original = middle.map((r) => r.text).join("");
  const next = fn(original);
  let i = 0;
  const out =
    next.length === original.length
      ? middle.map((r) => {
          const text = next.slice(i, i + r.text.length);
          i += r.text.length;
          return { text, style: r.style };
        })
      : middle;
  return normalizeRuns([...head, ...out, ...tail], fallbackStyle(runs, start));
}

// ---------------------------------------------------------------------------
// applyOp

const identity = (p: DocPosition): DocPosition => p;

/** Content ops accept ANY paragraph (top-level or table cell). */
function mustLocate(doc: Document, blockId: string): { loc: ParaLocation; block: Paragraph } {
  const loc = locateParagraph(doc, blockId);
  if (!loc) throw new Error(`paragraph ${blockId} not found`);
  return { loc, block: paragraphAt(doc, loc) };
}

// ---- container plumbing: structural ops work on the body OR a band story ----

export function containerBlocks(doc: Document, where: Container): Block[] {
  return where === "body" ? doc.blocks : (doc.section[where] ?? []);
}

function withContainerBlocks(doc: Document, where: Container, blocks: Block[]): Document {
  if (where === "body") return { ...doc, blocks };
  return { ...doc, section: { ...doc.section, [where]: blocks } };
}

/** Which container holds this TOP-LEVEL block (paragraphs in table cells are
 *  not top-level and return null). */
export function containerOf(doc: Document, blockId: string): { where: Container; index: number } | null {
  for (const where of ["body", ...BAND_CONTAINERS] as Container[]) {
    const index = containerBlocks(doc, where).findIndex((b) => b.id === blockId);
    if (index >= 0) return { where, index };
  }
  return null;
}

function mustTable(doc: Document, blockId: string): { where: Container; bi: number; block: TableBlock } {
  const found = containerOf(doc, blockId);
  const block = found ? containerBlocks(doc, found.where)[found.index] : undefined;
  if (!found || !block || block.kind !== "table") throw new Error(`table ${blockId} not found`);
  return { where: found.where, bi: found.index, block };
}

function replaceTable(doc: Document, where: Container, bi: number, table: TableBlock): Document {
  const blocks = containerBlocks(doc, where).slice();
  blocks[bi] = { ...table, revision: table.revision + 1 };
  return withContainerBlocks(doc, where, blocks);
}

/** Path-clone a table (body or band) replacing one cell's block list
 *  (revision bumped). */
function replaceCellBlocks(
  doc: Document,
  where: Container,
  bi: number,
  ri: number,
  ci: number,
  cellBlocks: Block[],
): Document {
  const blocks = containerBlocks(doc, where).slice();
  const table = blocks[bi] as TableBlock;
  const rows = table.rows.slice();
  const row = { cells: rows[ri]!.cells.slice() };
  row.cells[ci] = { ...row.cells[ci]!, blocks: cellBlocks };
  rows[ri] = row;
  blocks[bi] = { ...table, rows, revision: table.revision + 1 };
  return withContainerBlocks(doc, where, blocks);
}

/** First caret-capable paragraph in a set of rows (cells may start with images). */
export function firstParagraphInRows(rows: TableRow[]): Paragraph | undefined {
  for (const row of rows) {
    for (const cell of row.cells) {
      for (const b of cell.blocks) if (b.kind === "paragraph") return b;
    }
  }
  return undefined;
}

/** Find an image block anywhere editable: top-level containers or table cells. */
export type ImageLocation =
  | { kind: "top"; where: Container; index: number; image: ImageBlock }
  | { kind: "cell"; bi: number; ri: number; ci: number; ii: number; image: ImageBlock };

export function locateImage(doc: Document, blockId: string): ImageLocation | null {
  const found = containerOf(doc, blockId);
  if (found) {
    const block = containerBlocks(doc, found.where)[found.index];
    if (block?.kind === "image") return { kind: "top", where: found.where, index: found.index, image: block };
    return null;
  }
  for (let bi = 0; bi < doc.blocks.length; bi++) {
    const b = doc.blocks[bi]!;
    if (b.kind !== "table") continue;
    for (let ri = 0; ri < b.rows.length; ri++) {
      const row = b.rows[ri]!;
      for (let ci = 0; ci < row.cells.length; ci++) {
        const cell = row.cells[ci]!;
        for (let ii = 0; ii < cell.blocks.length; ii++) {
          const cb = cell.blocks[ii]!;
          if (cb.kind === "image" && cb.id === blockId) return { kind: "cell", bi, ri, ci, ii, image: cb };
        }
      }
    }
  }
  return null;
}

/** Where a block lives: top-level in a container (body or a header/footer band)
 *  or one level deep in a table cell. The path carries the container so callers
 *  can clone the right story. */
export type BlockLocation<B extends Block> =
  | { kind: "top"; where: Container; index: number; block: B }
  | { kind: "cell"; where: Container; bi: number; ri: number; ci: number; ii: number; block: B };

/** Find a block of a given kind anywhere editable — a top-level block of any
 *  container (body or a header/footer band) OR a block nested one level deep in
 *  a table cell (of a body OR band table). The container-aware generalization of
 *  locateImage; equation editing/alignment/deletion resolve through it so display
 *  equations imported into cells or bands stay live. Returns null when no block
 *  with that id and kind exists. */
export function locateBlock<K extends Block["kind"]>(
  doc: Document,
  blockId: string,
  blockKind: K,
): BlockLocation<Extract<Block, { kind: K }>> | null {
  type B = Extract<Block, { kind: K }>;
  const found = containerOf(doc, blockId);
  if (found) {
    const block = containerBlocks(doc, found.where)[found.index];
    if (block?.kind === blockKind) {
      return { kind: "top", where: found.where, index: found.index, block: block as B };
    }
    return null; // id is a top-level block, but not of the requested kind
  }
  // Not top-level — scan table cells in the body and every band story.
  for (const where of ["body", ...BAND_CONTAINERS] as Container[]) {
    const blocks = containerBlocks(doc, where);
    for (let bi = 0; bi < blocks.length; bi++) {
      const b = blocks[bi]!;
      if (b.kind !== "table") continue;
      for (let ri = 0; ri < b.rows.length; ri++) {
        const row = b.rows[ri]!;
        for (let ci = 0; ci < row.cells.length; ci++) {
          const cell = row.cells[ci]!;
          for (let ii = 0; ii < cell.blocks.length; ii++) {
            const cb = cell.blocks[ii]!;
            if (cb.kind === blockKind && cb.id === blockId) {
              return { kind: "cell", where, bi, ri, ci, ii, block: cb as B };
            }
          }
        }
      }
    }
  }
  return null;
}

/** Locate a display-equation block anywhere editable (body / band / table cell). */
export function locateEquation(doc: Document, blockId: string): BlockLocation<EquationBlock> | null {
  return locateBlock(doc, blockId, "equation");
}

/** Path-clone the document replacing a located block in place (top-level in any
 *  container, or one level deep in a table cell). */
function replaceLocatedBlock(doc: Document, loc: BlockLocation<Block>, block: Block): Document {
  if (loc.kind === "top") {
    const blocks = containerBlocks(doc, loc.where).slice();
    blocks[loc.index] = block;
    return withContainerBlocks(doc, loc.where, blocks);
  }
  const table = containerBlocks(doc, loc.where)[loc.bi] as TableBlock;
  const cellBlocks = table.rows[loc.ri]!.cells[loc.ci]!.blocks.slice();
  cellBlocks[loc.ii] = block;
  return replaceCellBlocks(doc, loc.where, loc.bi, loc.ri, loc.ci, cellBlocks);
}

/** The table's true grid-column count: the widest row measured in GRID columns,
 *  not cells — colSpan widens a cell and a rowSpan from an earlier row leaves a
 *  hole that this row's cells shift past (HTML table model). Counting cells
 *  alone undercounts any row with a colSpan or a vertical-merge hole. */
export function gridColumnCount(t: TableBlock): number {
  let maxCols = 1;
  const rowsRemaining: number[] = [];
  for (const row of t.rows) {
    let col = 0;
    for (const cell of row.cells) {
      while ((rowsRemaining[col] ?? 0) > 0) col++;
      const span = Math.max(1, cell.colSpan ?? 1);
      const rowSpan = Math.max(1, cell.rowSpan ?? 1);
      if (rowSpan > 1) for (let k = 0; k < span; k++) rowsRemaining[col + k] = rowSpan;
      col += span;
    }
    maxCols = Math.max(maxCols, col);
    for (let c = 0; c < rowsRemaining.length; c++) if (rowsRemaining[c]! > 0) rowsRemaining[c]!--;
  }
  return maxCols;
}

/** Column fractions normalized to the table's column count. */
export function effectiveFractions(t: TableBlock): number[] {
  const n = gridColumnCount(t);
  if (t.colFractions && t.colFractions.length === n) return t.colFractions;
  return Array.from({ length: n }, () => 1 / n);
}

function bump(block: Paragraph, runs: Run[]): Paragraph {
  return { ...block, runs, revision: block.revision + 1 };
}

export function applyOp(doc: Document, op: Op): ApplyResult {
  switch (op.type) {
    case "insertText": {
      const { loc, block } = mustLocate(doc, op.at.blockId);
      const runs = insertTextInRuns(block.runs, op.at.offset, op.text, op.style);
      const len = op.text.length;
      return {
        doc: replaceParagraphAt(doc, loc, bump(block, runs)),
        inverse: { type: "deleteRange", blockId: block.id, start: op.at.offset, end: op.at.offset + len },
        mapPosition: (p) =>
          p.blockId === block.id && p.offset >= op.at.offset
            ? { blockId: p.blockId, offset: p.offset + len }
            : p,
        dirtyBlockIds: [block.id],
      };
    }

    case "insertRuns": {
      const { loc, block } = mustLocate(doc, op.at.blockId);
      const len = textOfRuns(op.runs).length;
      const runs = insertRunsInRuns(block.runs, op.at.offset, op.runs);
      return {
        doc: replaceParagraphAt(doc, loc, bump(block, runs)),
        inverse: { type: "deleteRange", blockId: block.id, start: op.at.offset, end: op.at.offset + len },
        mapPosition: (p) =>
          p.blockId === block.id && p.offset >= op.at.offset
            ? { blockId: p.blockId, offset: p.offset + len }
            : p,
        dirtyBlockIds: [block.id],
      };
    }

    case "deleteRange": {
      const { loc, block } = mustLocate(doc, op.blockId);
      const removed = sliceRuns(block.runs, op.start, op.end);
      const runs = deleteInRuns(block.runs, op.start, op.end);
      const len = op.end - op.start;
      return {
        doc: replaceParagraphAt(doc, loc, bump(block, runs)),
        inverse: { type: "insertRuns", at: { blockId: block.id, offset: op.start }, runs: removed },
        mapPosition: (p) => {
          if (p.blockId !== block.id || p.offset <= op.start) return p;
          return { blockId: p.blockId, offset: p.offset >= op.end ? p.offset - len : op.start };
        },
        dirtyBlockIds: [block.id],
      };
    }

    case "setRuns": {
      const { loc, block } = mustLocate(doc, op.blockId);
      return {
        doc: replaceParagraphAt(doc, loc, bump(block, op.runs)),
        inverse: { type: "setRuns", blockId: block.id, runs: block.runs },
        mapPosition: identity, // callers use setRuns for length-preserving restyles
        dirtyBlockIds: [block.id],
      };
    }

    case "setParaStyle": {
      const { loc, block } = mustLocate(doc, op.blockId);
      const oldPatch: Partial<ParaStyle> = {};
      for (const key of Object.keys(op.patch) as (keyof ParaStyle)[]) {
        // @ts-expect-error — keyed copy of the previous values for the inverse
        oldPatch[key] = block.style[key];
      }
      const styled: Paragraph = {
        ...block,
        style: { ...block.style, ...op.patch },
        revision: block.revision + 1,
      };
      return {
        doc: replaceParagraphAt(doc, loc, styled),
        inverse: { type: "setParaStyle", blockId: block.id, patch: oldPatch },
        mapPosition: identity,
        dirtyBlockIds: [block.id],
      };
    }

    case "splitParagraph": {
      const loc = locateParagraph(doc, op.at.blockId);
      if (!loc) throw new Error(`paragraph ${op.at.blockId} not found`);
      const block = paragraphAt(doc, loc);
      const [headRuns, tailRuns] = splitRunsAt(block.runs, op.at.offset);
      const carry = fallbackStyle(block.runs, op.at.offset);
      const head = bump(block, normalizeRuns(headRuns, carry));
      const tail: Paragraph = {
        kind: "paragraph",
        id: op.newBlockId,
        revision: 0,
        runs: normalizeRuns(tailRuns, carry),
        style: op.newStyle ?? { ...block.style },
        // The tail's block-level content control is declared by the op (the split
        // command carries the source paragraph's path; the merge inverse restores
        // the original tail's). Absent → the tail belongs to no block control.
        ...(op.newSdtPath?.length ? { sdtPath: op.newSdtPath } : {}),
      };
      let next: Document;
      if (loc.kind === "cell") {
        const table = containerBlocks(doc, loc.where)[loc.bi] as TableBlock;
        const cellBlocks = table.rows[loc.ri]!.cells[loc.ci]!.blocks.slice();
        cellBlocks.splice(loc.pi, 1, head, tail);
        next = replaceCellBlocks(doc, loc.where, loc.bi, loc.ri, loc.ci, cellBlocks);
      } else if (loc.kind === "footnote") {
        const paras = doc.footnotes![loc.noteId]!.slice();
        paras.splice(loc.pi, 1, head, tail);
        next = { ...doc, footnotes: { ...doc.footnotes, [loc.noteId]: paras } };
      } else if (loc.kind === "endnote") {
        const paras = doc.endnotes![loc.noteId]!.slice();
        paras.splice(loc.pi, 1, head, tail);
        next = { ...doc, endnotes: { ...doc.endnotes, [loc.noteId]: paras } };
      } else {
        const where: Container = loc.kind === "band" ? loc.band : "body";
        const blocks = containerBlocks(doc, where).slice();
        blocks.splice(loc.bi, 1, head, tail);
        next = withContainerBlocks(doc, where, blocks);
      }
      return {
        doc: next,
        inverse: { type: "mergeParagraphs", firstBlockId: block.id },
        mapPosition: (p) =>
          p.blockId === block.id && p.offset >= op.at.offset
            ? { blockId: op.newBlockId, offset: p.offset - op.at.offset }
            : p,
        dirtyBlockIds: [block.id, op.newBlockId],
      };
    }

    case "mergeParagraphs": {
      const loc = locateParagraph(doc, op.firstBlockId);
      if (!loc) throw new Error(`paragraph ${op.firstBlockId} not found`);
      const block = paragraphAt(doc, loc);
      let nextPara: Paragraph;
      if (loc.kind === "cell") {
        const table = containerBlocks(doc, loc.where)[loc.bi] as TableBlock;
        const candidate = table.rows[loc.ri]!.cells[loc.ci]!.blocks[loc.pi + 1];
        if (!candidate || candidate.kind !== "paragraph") {
          throw new Error("mergeParagraphs: no next paragraph in cell");
        }
        nextPara = candidate;
      } else if (loc.kind === "footnote") {
        const candidate = doc.footnotes![loc.noteId]![loc.pi + 1];
        if (!candidate) throw new Error("mergeParagraphs: no next paragraph in footnote");
        nextPara = candidate;
      } else if (loc.kind === "endnote") {
        const candidate = doc.endnotes![loc.noteId]![loc.pi + 1];
        if (!candidate) throw new Error("mergeParagraphs: no next paragraph in endnote");
        nextPara = candidate;
      } else {
        const where: Container = loc.kind === "band" ? loc.band : "body";
        const candidate = containerBlocks(doc, where)[loc.bi + 1];
        if (!candidate || candidate.kind !== "paragraph") {
          throw new Error("mergeParagraphs: no next paragraph");
        }
        nextPara = candidate;
      }
      const headLen = textOfRuns(block.runs).length;
      const merged = bump(
        block,
        normalizeRuns([...block.runs, ...nextPara.runs], fallbackStyle(block.runs, headLen)),
      );
      let next: Document;
      if (loc.kind === "cell") {
        const table = containerBlocks(doc, loc.where)[loc.bi] as TableBlock;
        const cellBlocks = table.rows[loc.ri]!.cells[loc.ci]!.blocks.slice();
        cellBlocks.splice(loc.pi, 2, merged);
        next = replaceCellBlocks(doc, loc.where, loc.bi, loc.ri, loc.ci, cellBlocks);
      } else if (loc.kind === "footnote") {
        const paras = doc.footnotes![loc.noteId]!.slice();
        paras.splice(loc.pi, 2, merged);
        next = { ...doc, footnotes: { ...doc.footnotes, [loc.noteId]: paras } };
      } else if (loc.kind === "endnote") {
        const paras = doc.endnotes![loc.noteId]!.slice();
        paras.splice(loc.pi, 2, merged);
        next = { ...doc, endnotes: { ...doc.endnotes, [loc.noteId]: paras } };
      } else {
        const where: Container = loc.kind === "band" ? loc.band : "body";
        const blocks = containerBlocks(doc, where).slice();
        blocks.splice(loc.bi, 2, merged);
        next = withContainerBlocks(doc, where, blocks);
      }
      return {
        doc: next,
        inverse: {
          type: "splitParagraph",
          at: { blockId: block.id, offset: headLen },
          newBlockId: nextPara.id,
          newStyle: nextPara.style,
          // Restore the merged-away tail's block-level control on undo.
          ...(nextPara.sdtPath?.length ? { newSdtPath: nextPara.sdtPath } : {}),
        },
        mapPosition: (p) =>
          p.blockId === nextPara.id ? { blockId: block.id, offset: headLen + p.offset } : p,
        dirtyBlockIds: [block.id, nextPara.id],
      };
    }

    case "insertBlock": {
      const where = op.where ?? "body";
      const blocks = containerBlocks(doc, where).slice();
      blocks.splice(op.index, 0, op.block);
      return {
        doc: withContainerBlocks(doc, where, blocks),
        inverse: { type: "removeBlock", blockId: op.block.id },
        mapPosition: identity,
        dirtyBlockIds: [op.block.id],
      };
    }

    case "removeBlock": {
      const found = containerOf(doc, op.blockId);
      if (!found) throw new Error(`block ${op.blockId} not found`);
      const blocks = containerBlocks(doc, found.where).slice();
      const block = blocks[found.index]!;
      blocks.splice(found.index, 1);
      const neighbor = blocks[Math.min(found.index, blocks.length - 1)];
      return {
        doc: withContainerBlocks(doc, found.where, blocks),
        inverse: { type: "insertBlock", index: found.index, block, where: found.where },
        mapPosition: (p) =>
          p.blockId === op.blockId && neighbor
            ? { blockId: neighbor.id, offset: 0 }
            : p,
        dirtyBlockIds: [op.blockId],
      };
    }

    case "setImageProps": {
      const loc = locateImage(doc, op.blockId);
      if (!loc) throw new Error(`image ${op.blockId} not found`);
      const block = loc.image;
      // Inverse restores prior values; a field absent before is cleared (null)
      // on undo. wrap/anchor use null to mean "clear" (they're exclusive states).
      const oldPatch: ImagePropsPatch = {};
      if (op.patch.widthPx !== undefined) oldPatch.widthPx = block.widthPx;
      if (op.patch.heightPx !== undefined) oldPatch.heightPx = block.heightPx;
      if (op.patch.align !== undefined) oldPatch.align = block.align;
      if (op.patch.wrap !== undefined) oldPatch.wrap = block.wrap ?? null;
      if (op.patch.anchor !== undefined) oldPatch.anchor = block.anchor ?? null;
      if (op.patch.crop !== undefined) oldPatch.crop = block.crop ?? null;
      const updated: ImageBlock = { ...block, revision: block.revision + 1 };
      if (op.patch.widthPx !== undefined) updated.widthPx = op.patch.widthPx;
      if (op.patch.heightPx !== undefined) updated.heightPx = op.patch.heightPx;
      if (op.patch.align !== undefined) updated.align = op.patch.align;
      if (op.patch.wrap !== undefined) {
        if (op.patch.wrap === null) delete updated.wrap;
        else updated.wrap = op.patch.wrap;
      }
      if (op.patch.anchor !== undefined) {
        if (op.patch.anchor === null) delete updated.anchor;
        else updated.anchor = op.patch.anchor;
      }
      if (op.patch.crop !== undefined) {
        if (op.patch.crop === null) delete updated.crop;
        else updated.crop = normalizeCrop(op.patch.crop);
      }
      let next: Document;
      if (loc.kind === "top") {
        const blocks = containerBlocks(doc, loc.where).slice();
        blocks[loc.index] = updated;
        next = withContainerBlocks(doc, loc.where, blocks);
      } else {
        const table = doc.blocks[loc.bi] as TableBlock;
        const cellBlocks = table.rows[loc.ri]!.cells[loc.ci]!.blocks.slice();
        cellBlocks[loc.ii] = updated;
        next = replaceCellBlocks(doc, "body", loc.bi, loc.ri, loc.ci, cellBlocks); // locateImage cells are body-only

      }
      return {
        doc: next,
        inverse: { type: "setImageProps", blockId: op.blockId, patch: oldPatch },
        mapPosition: identity,
        dirtyBlockIds: [op.blockId],
      };
    }

    case "setEquation": {
      // Replace a display equation's MathML (the editor's Apply). Container-aware:
      // the equation may live in the body, a header/footer band, or a table cell.
      const loc = locateEquation(doc, op.blockId);
      if (!loc) throw new Error(`equation ${op.blockId} not found`);
      const old = loc.block.equation;
      const updated: EquationBlock = { ...loc.block, revision: loc.block.revision + 1, equation: op.equation };
      return {
        doc: replaceLocatedBlock(doc, loc, updated),
        inverse: { type: "setEquation", blockId: op.blockId, equation: old },
        mapPosition: identity,
        dirtyBlockIds: [op.blockId],
      };
    }

    case "setEquationAlign": {
      const loc = locateEquation(doc, op.blockId);
      if (!loc) throw new Error(`equation ${op.blockId} not found`);
      const old = loc.block.align ?? "center";
      const updated: EquationBlock = { ...loc.block, revision: loc.block.revision + 1, align: op.align };
      return {
        doc: replaceLocatedBlock(doc, loc, updated),
        inverse: { type: "setEquationAlign", blockId: op.blockId, align: old },
        mapPosition: identity,
        dirtyBlockIds: [op.blockId],
      };
    }

    case "setTableColFractions": {
      const { where, bi, block } = mustTable(doc, op.blockId);
      const old = effectiveFractions(block);
      return {
        doc: replaceTable(doc, where, bi, { ...block, colFractions: op.fractions }),
        inverse: { type: "setTableColFractions", blockId: op.blockId, fractions: old },
        mapPosition: identity,
        dirtyBlockIds: [op.blockId],
      };
    }

    case "setTableWidthMode": {
      const { where, bi, block } = mustTable(doc, op.blockId);
      const old = block.widthMode;
      const next: TableBlock = { ...block };
      if (op.mode && op.mode !== "fixed") next.widthMode = op.mode;
      else delete next.widthMode;
      return {
        doc: replaceTable(doc, where, bi, next),
        inverse: { type: "setTableWidthMode", blockId: op.blockId, mode: old },
        mapPosition: identity,
        dirtyBlockIds: [op.blockId],
      };
    }

    case "setTablePreferredWidth": {
      const { where, bi, block } = mustTable(doc, op.blockId);
      const old = block.preferredWidth;
      const next: TableBlock = { ...block };
      if (op.width) next.preferredWidth = op.width;
      else delete next.preferredWidth;
      return {
        doc: replaceTable(doc, where, bi, next),
        inverse: { type: "setTablePreferredWidth", blockId: op.blockId, width: old ?? null },
        mapPosition: identity,
        dirtyBlockIds: [op.blockId],
      };
    }

    case "setTableAlign": {
      const { where, bi, block } = mustTable(doc, op.blockId);
      const old = block.align;
      const next: TableBlock = { ...block };
      if (op.align && op.align !== "left") next.align = op.align;
      else delete next.align;
      return {
        doc: replaceTable(doc, where, bi, next),
        inverse: { type: "setTableAlign", blockId: op.blockId, align: old ?? null },
        mapPosition: identity,
        dirtyBlockIds: [op.blockId],
      };
    }

    case "setTableProps": {
      // Table-LEVEL props (indent + cascade defaults). Patch keys present are
      // applied (a value sets, `null` clears); the inverse captures each touched
      // field's prior value (null = was absent) so undo restores it exactly.
      const { where, bi, block } = mustTable(doc, op.blockId);
      const keys = Object.keys(op.patch) as (keyof TablePropsPatch)[];
      const oldPatch: TablePropsPatch = {};
      const next: Record<string, unknown> = { ...block };
      const blockRec = block as unknown as Record<string, unknown>;
      const oldRec = oldPatch as Record<string, unknown>;
      for (const key of keys) {
        oldRec[key] = blockRec[key] ?? null;
        const val = op.patch[key];
        if (val === undefined || val === null) delete next[key];
        else next[key] = val;
      }
      return {
        doc: replaceTable(doc, where, bi, next as unknown as TableBlock),
        inverse: { type: "setTableProps", blockId: op.blockId, patch: oldPatch },
        mapPosition: identity,
        dirtyBlockIds: [op.blockId],
      };
    }

    case "setTableRow": {
      const { where, bi, block } = mustTable(doc, op.tableId);
      const old = block.rows[op.rowIndex];
      if (!old) throw new Error("setTableRow: no such row");
      const removedIds = new Set(old.cells.flatMap((c) => c.blocks.map((p) => p.id)));
      const rows = block.rows.slice();
      rows[op.rowIndex] = op.row;
      const fallback = firstParagraphInRows([op.row]);
      return {
        doc: replaceTable(doc, where, bi, { ...block, rows }),
        inverse: { type: "setTableRow", tableId: op.tableId, rowIndex: op.rowIndex, row: old },
        mapPosition: (p) => {
          if (!removedIds.has(p.blockId)) return p;
          // position survives if its paragraph still exists in the new row
          const kept = op.row.cells.some((c) => c.blocks.some((b) => b.id === p.blockId));
          return kept || !fallback ? p : { blockId: fallback.id, offset: 0 };
        },
        dirtyBlockIds: [op.tableId],
      };
    }

    case "insertTableRow": {
      const { where, bi, block } = mustTable(doc, op.tableId);
      const rows = block.rows.slice();
      rows.splice(op.rowIndex, 0, op.row);
      return {
        doc: replaceTable(doc, where, bi, { ...block, rows }),
        inverse: { type: "removeTableRow", tableId: op.tableId, rowIndex: op.rowIndex },
        mapPosition: identity,
        dirtyBlockIds: [op.tableId],
      };
    }

    case "removeTableRow": {
      const { where, bi, block } = mustTable(doc, op.tableId);
      const removed = block.rows[op.rowIndex];
      if (!removed) throw new Error("removeTableRow: no such row");
      const rows = block.rows.slice();
      rows.splice(op.rowIndex, 1);
      const removedIds = new Set(removed.cells.flatMap((c) => c.blocks.map((p) => p.id)));
      const fallbackRow = rows[Math.min(op.rowIndex, rows.length - 1)];
      const fallback = fallbackRow ? firstParagraphInRows([fallbackRow]) : undefined;
      return {
        doc: replaceTable(doc, where, bi, { ...block, rows }),
        inverse: { type: "insertTableRow", tableId: op.tableId, rowIndex: op.rowIndex, row: removed },
        mapPosition: (p) =>
          removedIds.has(p.blockId) && fallback ? { blockId: fallback.id, offset: 0 } : p,
        dirtyBlockIds: [op.tableId],
      };
    }

    case "setRowHeight": {
      // Interactive row-drag height (w:trHeight). Mirrors the column op: replace a
      // single row's `props.height`, leaving the rest of its props untouched, and
      // invert to the prior height (null = none) for a free undo.
      const { where, bi, block } = mustTable(doc, op.tableId);
      const row = block.rows[op.rowIndex];
      if (!row) throw new Error("setRowHeight: no such row");
      const old = row.props?.height ?? null;
      const nextRow: TableRow = { ...row };
      const props: RowProps = { ...(row.props ?? {}) };
      if (op.height) props.height = op.height;
      else delete props.height;
      if (Object.keys(props).length) nextRow.props = props;
      else delete nextRow.props;
      const rows = block.rows.slice();
      rows[op.rowIndex] = nextRow;
      return {
        doc: replaceTable(doc, where, bi, { ...block, rows }),
        inverse: { type: "setRowHeight", tableId: op.tableId, rowIndex: op.rowIndex, height: old },
        mapPosition: identity,
        dirtyBlockIds: [op.tableId],
      };
    }

    case "setListDefinition": {
      const old = doc.lists?.[op.listId] ?? null;
      const lists = { ...(doc.lists ?? {}) };
      if (op.def) lists[op.listId] = op.def;
      else delete lists[op.listId];
      return {
        doc: { ...doc, lists },
        inverse: { type: "setListDefinition", listId: op.listId, def: old },
        mapPosition: identity,
        // Indents come from the definition → every paragraph in the list must
        // re-measure; their (revision,width) line-cache keys change with width.
        dirtyBlockIds: [],
      };
    }

    case "setTableStructure": {
      const { where, bi, block } = mustTable(doc, op.tableId);
      const oldIds = new Set(
        block.rows.flatMap((r) => r.cells.flatMap((c) => c.blocks.map((p) => p.id))),
      );
      const fallback = firstParagraphInRows(op.rows);
      const next: TableBlock = { ...block, rows: op.rows };
      if (op.colFractions) next.colFractions = op.colFractions;
      else delete next.colFractions;
      const inverse: Op = { type: "setTableStructure", tableId: op.tableId, rows: block.rows };
      if (block.colFractions) inverse.colFractions = block.colFractions;
      return {
        doc: replaceTable(doc, where, bi, next),
        inverse,
        mapPosition: (p) => {
          if (!oldIds.has(p.blockId)) return p;
          const kept = op.rows.some((r) =>
            r.cells.some((c) => c.blocks.some((b) => b.id === p.blockId)),
          );
          return kept || !fallback ? p : { blockId: fallback.id, offset: 0 };
        },
        dirtyBlockIds: [op.tableId],
      };
    }

    case "insertTableColumn": {
      const { where, bi, block } = mustTable(doc, op.tableId);
      const snapshot: Op = { type: "setTableStructure", tableId: op.tableId, rows: block.rows };
      if (block.colFractions) snapshot.colFractions = block.colFractions;
      // Span-aware: a merged cell covering the insertion point grows by one
      // column instead of having a new cell slotted into its middle.
      const rows = block.rows.map((row, ri) => {
        const cells: TableCell[] = [];
        let col = 0;
        let inserted = false;
        for (const cell of row.cells) {
          const span = cell.colSpan ?? 1;
          if (!inserted && op.colIndex > col && op.colIndex < col + span) {
            cells.push({ ...cell, colSpan: span + 1 });
            inserted = true;
          } else if (!inserted && op.colIndex === col) {
            const fresh = op.cells[ri];
            if (fresh) cells.push(fresh);
            cells.push(cell);
            inserted = true;
          } else {
            cells.push(cell);
          }
          col += span;
        }
        if (!inserted) {
          const fresh = op.cells[ri];
          if (fresh) cells.push(fresh); // append at the right edge
        }
        return { cells };
      });
      let fractions = op.fractions;
      if (!fractions) {
        const old = effectiveFractions(block);
        const fresh = 1 / (old.length + 1);
        fractions = old.map((f) => f * (1 - fresh));
        fractions.splice(op.colIndex, 0, fresh);
      }
      return {
        doc: replaceTable(doc, where, bi, { ...block, rows, colFractions: fractions }),
        inverse: snapshot,
        mapPosition: identity,
        dirtyBlockIds: [op.tableId],
      };
    }

    case "setStylesheet": {
      const old = doc.stylesheet ?? { styles: [], defaultStyleId: "Normal" };
      return {
        doc: { ...doc, stylesheet: op.stylesheet },
        inverse: { type: "setStylesheet", stylesheet: old },
        mapPosition: identity,
        dirtyBlockIds: [], // restyling paragraphs happens via setRuns/setParaStyle ops
      };
    }

    case "setTableStyleSheet": {
      const old = doc.tableStyles ?? {};
      return {
        doc: { ...doc, tableStyles: op.tableStyles },
        inverse: { type: "setTableStyleSheet", tableStyles: old },
        mapPosition: identity,
        dirtyBlockIds: [], // re-baking cells happens via the table-replace ops
      };
    }

    case "setTableStyleRef": {
      const { where, bi, block } = mustTable(doc, op.tableId);
      const oldStyleId = block.styleId ?? null;
      const oldCond = block.condOverrides ?? null;
      const next: TableBlock = { ...block };
      if (op.styleId === null) delete next.styleId;
      else next.styleId = op.styleId;
      if (op.condOverrides !== undefined) {
        if (op.condOverrides === null) delete next.condOverrides;
        else next.condOverrides = op.condOverrides;
      }
      return {
        doc: replaceTable(doc, where, bi, next),
        inverse: { type: "setTableStyleRef", tableId: op.tableId, styleId: oldStyleId, condOverrides: oldCond },
        mapPosition: identity,
        dirtyBlockIds: [op.tableId],
      };
    }

    case "setSectionProps": {
      const s = doc.section;
      const old: SectionGeometry = {
        pageWidthPx: s.pageWidthPx,
        pageHeightPx: s.pageHeightPx,
        marginPx: { ...s.marginPx },
        columns: s.columns
          ? { ...s.columns, ...(s.columns.cols ? { cols: s.columns.cols.map((c) => ({ ...c })) } : {}) }
          : null,
        pageNumberStart: s.pageNumberStart ?? null,
        headerDistancePx: s.headerDistancePx ?? null,
        footerDistancePx: s.footerDistancePx ?? null,
        pageColorHex: s.pageColorHex ?? null,
        pageBorders: s.pageBorders ?? null,
        breakType: s.breakType ?? null,
        lineNumbering: s.lineNumbering ? { ...s.lineNumbering } : null,
      };
      const next = {
        ...s,
        pageWidthPx: op.geometry.pageWidthPx,
        pageHeightPx: op.geometry.pageHeightPx,
        marginPx: op.geometry.marginPx,
      };
      if (op.geometry.columns) next.columns = op.geometry.columns;
      else delete next.columns;
      if (op.geometry.pageNumberStart !== null) next.pageNumberStart = op.geometry.pageNumberStart;
      else delete next.pageNumberStart;
      if (op.geometry.headerDistancePx !== null) next.headerDistancePx = op.geometry.headerDistancePx;
      else delete next.headerDistancePx;
      if (op.geometry.footerDistancePx !== null) next.footerDistancePx = op.geometry.footerDistancePx;
      else delete next.footerDistancePx;
      if (op.geometry.pageColorHex !== null) next.pageColorHex = op.geometry.pageColorHex;
      else delete next.pageColorHex;
      if (op.geometry.pageBorders !== null) next.pageBorders = op.geometry.pageBorders;
      else delete next.pageBorders;
      if (op.geometry.breakType !== null) next.breakType = op.geometry.breakType;
      else delete next.breakType;
      if (op.geometry.lineNumbering !== null) next.lineNumbering = op.geometry.lineNumbering;
      else delete next.lineNumbering;
      return {
        doc: { ...doc, section: next },
        inverse: { type: "setSectionProps", geometry: old },
        mapPosition: identity,
        dirtyBlockIds: [], // geometry change → full re-walk; line caches keyed by width
      };
    }

    case "setFootnote": {
      const old = doc.footnotes?.[op.noteId] ?? null;
      const footnotes = { ...(doc.footnotes ?? {}) };
      if (op.paras) footnotes[op.noteId] = op.paras;
      else delete footnotes[op.noteId];
      const removedIds = new Set((old ?? []).map((p) => p.id));
      const fallback = doc.blocks.find((b) => b.kind === "paragraph");
      return {
        doc: { ...doc, footnotes },
        inverse: { type: "setFootnote", noteId: op.noteId, paras: old },
        mapPosition: (p) => {
          if (!removedIds.has(p.blockId)) return p;
          const kept = (op.paras ?? []).some((b) => b.id === p.blockId);
          return kept || !fallback ? p : { blockId: fallback.id, offset: 0 };
        },
        dirtyBlockIds: [],
      };
    }

    case "setEndnote": {
      const old = doc.endnotes?.[op.noteId] ?? null;
      const endnotes = { ...(doc.endnotes ?? {}) };
      if (op.paras) endnotes[op.noteId] = op.paras;
      else delete endnotes[op.noteId];
      const removedIds = new Set((old ?? []).map((p) => p.id));
      const fallback = doc.blocks.find((b) => b.kind === "paragraph");
      return {
        doc: { ...doc, endnotes },
        inverse: { type: "setEndnote", noteId: op.noteId, paras: old },
        mapPosition: (p) => {
          if (!removedIds.has(p.blockId)) return p;
          const kept = (op.paras ?? []).some((b) => b.id === p.blockId);
          return kept || !fallback ? p : { blockId: fallback.id, offset: 0 };
        },
        dirtyBlockIds: [],
      };
    }

    case "setSdtProps": {
      const old = doc.sdts?.[op.id] ?? null;
      const sdts = { ...(doc.sdts ?? {}) };
      if (op.props) sdts[op.id] = op.props;
      else delete sdts[op.id];
      return {
        doc: { ...doc, sdts },
        inverse: { type: "setSdtProps", id: op.id, props: old },
        mapPosition: identity,
        dirtyBlockIds: [], // run markers change via setRuns/applyStylePatch ops
      };
    }

    case "setField": {
      const old = doc.fields?.[op.id] ?? null;
      const fields = { ...(doc.fields ?? {}) };
      if (op.def) fields[op.id] = op.def;
      else delete fields[op.id];
      return {
        doc: { ...doc, fields },
        inverse: { type: "setField", id: op.id, def: old },
        mapPosition: identity,
        dirtyBlockIds: [], // run/block fieldId markers change via setRuns/insert ops
      };
    }

    case "setTocInstruction": {
      const old = doc.tocInstruction ?? null;
      const next = { ...doc };
      if (op.instruction !== null) next.tocInstruction = op.instruction;
      else delete next.tocInstruction;
      return {
        doc: next,
        inverse: { type: "setTocInstruction", instruction: old },
        mapPosition: identity,
        dirtyBlockIds: [], // entries are regenerated via insert/removeBlock ops in the same tx
      };
    }

    case "setBookmark": {
      const old = doc.bookmarks?.[op.name] ?? null;
      const bookmarks = { ...(doc.bookmarks ?? {}) };
      if (op.range) bookmarks[op.name] = op.range;
      else delete bookmarks[op.name];
      return {
        doc: { ...doc, bookmarks },
        inverse: { type: "setBookmark", name: op.name, range: old },
        mapPosition: identity,
        dirtyBlockIds: [],
      };
    }

    case "setSectionBand": {
      const old = doc.section[op.band] ?? null;
      const section = { ...doc.section };
      if (op.blocks) section[op.band] = op.blocks;
      else delete section[op.band];
      const removedIds = new Set(
        (old ?? []).filter((b) => b.kind === "paragraph").map((b) => b.id),
      );
      const fallback = doc.blocks.find((b) => b.kind === "paragraph");
      return {
        doc: { ...doc, section },
        inverse: { type: "setSectionBand", band: op.band, blocks: old },
        mapPosition: (p) => {
          if (!removedIds.has(p.blockId)) return p;
          const kept = (op.blocks ?? []).some((b) => b.id === p.blockId);
          return kept || !fallback ? p : { blockId: fallback.id, offset: 0 };
        },
        dirtyBlockIds: [],
      };
    }

    case "removeTableColumn": {
      const { where, bi, block } = mustTable(doc, op.tableId);
      const snapshot: Op = { type: "setTableStructure", tableId: op.tableId, rows: block.rows };
      if (block.colFractions) snapshot.colFractions = block.colFractions;
      const oldFractions = effectiveFractions(block);
      const removedIds = new Set<string>();
      // Span-aware: a merged cell covering the removed column shrinks; an
      // unmerged cell at that column is dropped.
      const rows = block.rows.map((row) => {
        const cells: TableCell[] = [];
        let col = 0;
        for (const cell of row.cells) {
          const span = cell.colSpan ?? 1;
          if (op.colIndex >= col && op.colIndex < col + span) {
            if (span > 1) {
              const shrunk: TableCell = { ...cell };
              if (span - 1 > 1) shrunk.colSpan = span - 1;
              else delete shrunk.colSpan;
              cells.push(shrunk);
            } else {
              for (const p of cell.blocks) removedIds.add(p.id);
            }
          } else {
            cells.push(cell);
          }
          col += span;
        }
        return { cells };
      });
      const rest = oldFractions.filter((_, i) => i !== op.colIndex);
      const sum = rest.reduce((s, f) => s + f, 0) || 1;
      const fractions = rest.map((f) => f / sum);
      const fallback = firstParagraphInRows(rows);
      return {
        doc: replaceTable(doc, where, bi, { ...block, rows, colFractions: fractions }),
        inverse: snapshot,
        mapPosition: (p) =>
          removedIds.has(p.blockId) && fallback ? { blockId: fallback.id, offset: 0 } : p,
        dirtyBlockIds: [op.tableId],
      };
    }
  }
}
