// Shared text helpers over the model: paragraph text access, grapheme/word
// boundaries (Intl.Segmenter — the same segmentation pretext breaks lines on),
// and style-at-position for Word-style inheritance when typing.

import type { BandContainer, Block, CharStyle, Document, Paragraph, Run } from "./document";
import { BAND_CONTAINERS } from "./document";

export const graphemes = new Intl.Segmenter(undefined, { granularity: "grapheme" });
export const words = new Intl.Segmenter(undefined, { granularity: "word" });

const paragraphsInBlocks = (blocks: Block[], includeCells: boolean): Paragraph[] => {
  const out: Paragraph[] = [];
  for (const b of blocks) {
    if (b.kind === "paragraph") out.push(b);
    else if (b.kind === "table" && includeCells) {
      for (const row of b.rows) {
        for (const cell of row.cells) out.push(...paragraphsInBlocks(cell.blocks, includeCells));
      }
    }
  }
  return out;
};

/** The two band STORIES from the UI's point of view (story mode is entered per
 *  header/footer area; which variant CONTAINER that resolves to depends on the
 *  page — see Page.headerSource/footerSource). */
export type BandName = "header" | "footer";

/** Editable paragraphs of a margin band container, document order — INCLUDING
 *  band-table cells (imported footers are routinely tables holding text next
 *  to a page-number paragraph). */
export const bandParagraphs = (doc: Document, band: BandContainer): Paragraph[] =>
  paragraphsInBlocks(doc.section[band] ?? [], true);

/** All editable paragraphs in document order: body (including table cells),
 *  then every band story (header/footer + first/even variants). This is the
 *  index space commands use. */
export const paragraphsOf = (doc: Document): Paragraph[] => [
  ...paragraphsInBlocks(doc.blocks, true),
  ...BAND_CONTAINERS.flatMap((band) => bandParagraphs(doc, band)),
  ...Object.values(doc.footnotes ?? {}).flat(),
];

export const blockById = (doc: Document, blockId: string): Paragraph | undefined =>
  paragraphsOf(doc).find((b) => b.id === blockId);

export const blockIndexOf = (doc: Document, blockId: string): number =>
  paragraphsOf(doc).findIndex((b) => b.id === blockId);

// ---------------------------------------------------------------------------
// Paragraph locator — content ops work on any paragraph, wherever it lives.

export type ParaLocation =
  | { kind: "top"; bi: number }
  | { kind: "cell"; where: "body" | BandContainer; bi: number; ri: number; ci: number; pi: number }
  | { kind: "band"; band: BandContainer; bi: number }
  | { kind: "footnote"; noteId: string; pi: number };

/** Top-level block list of a container ("body" or a band story). */
export const containerListOf = (doc: Document, where: "body" | BandContainer): Block[] =>
  where === "body" ? doc.blocks : (doc.section[where] ?? []);

function locateInBlocks(
  blocks: Block[],
  where: "body" | BandContainer,
  blockId: string,
): ParaLocation | null {
  for (let bi = 0; bi < blocks.length; bi++) {
    const b = blocks[bi]!;
    if (b.kind === "paragraph") {
      if (b.id === blockId) {
        return where === "body" ? { kind: "top", bi } : { kind: "band", band: where, bi };
      }
    } else if (b.kind === "table") {
      for (let ri = 0; ri < b.rows.length; ri++) {
        const row = b.rows[ri]!;
        for (let ci = 0; ci < row.cells.length; ci++) {
          const cell = row.cells[ci]!;
          for (let pi = 0; pi < cell.blocks.length; pi++) {
            const cb = cell.blocks[pi]!;
            // One level deep: paragraphs of nested tables stay read-only.
            if (cb.kind === "paragraph" && cb.id === blockId) {
              return { kind: "cell", where, bi, ri, ci, pi };
            }
          }
        }
      }
    }
  }
  return null;
}

export function locateParagraph(doc: Document, blockId: string): ParaLocation | null {
  const body = locateInBlocks(doc.blocks, "body", blockId);
  if (body) return body;
  for (const band of BAND_CONTAINERS) {
    const hit = locateInBlocks(doc.section[band] ?? [], band, blockId);
    if (hit) return hit;
  }
  for (const [noteId, paras] of Object.entries(doc.footnotes ?? {})) {
    const pi = paras.findIndex((p) => p.id === blockId);
    if (pi >= 0) return { kind: "footnote", noteId, pi };
  }
  return null;
}

export function paragraphAt(doc: Document, loc: ParaLocation): Paragraph {
  if (loc.kind === "top") return doc.blocks[loc.bi] as Paragraph;
  if (loc.kind === "band") return doc.section[loc.band]![loc.bi] as Paragraph;
  if (loc.kind === "footnote") return doc.footnotes![loc.noteId]![loc.pi]!;
  const table = containerListOf(doc, loc.where)[loc.bi] as Extract<Block, { kind: "table" }>;
  return table.rows[loc.ri]!.cells[loc.ci]!.blocks[loc.pi] as Paragraph;
}

/** Immutable path-clone replace; bumps the containing table's revision too. */
export function replaceParagraphAt(doc: Document, loc: ParaLocation, p: Paragraph): Document {
  if (loc.kind === "footnote") {
    const paras = doc.footnotes![loc.noteId]!.slice();
    paras[loc.pi] = p;
    return { ...doc, footnotes: { ...doc.footnotes, [loc.noteId]: paras } };
  }
  if (loc.kind === "band") {
    const blocks = (doc.section[loc.band] ?? []).slice();
    blocks[loc.bi] = p;
    return { ...doc, section: { ...doc.section, [loc.band]: blocks } };
  }
  if (loc.kind === "top") {
    const blocks = doc.blocks.slice();
    blocks[loc.bi] = p;
    return { ...doc, blocks };
  }
  const blocks = containerListOf(doc, loc.where).slice();
  const table = blocks[loc.bi] as Extract<Block, { kind: "table" }>;
  const rows = table.rows.slice();
  const row = { cells: rows[loc.ri]!.cells.slice() };
  const cell = { ...row.cells[loc.ci]!, blocks: row.cells[loc.ci]!.blocks.slice() };
  cell.blocks[loc.pi] = p;
  row.cells[loc.ci] = cell;
  rows[loc.ri] = row;
  blocks[loc.bi] = { ...table, rows, revision: table.revision + 1 };
  if (loc.where === "body") return { ...doc, blocks };
  return { ...doc, section: { ...doc.section, [loc.where]: blocks } };
}

export const isInCell = (doc: Document, blockId: string): boolean =>
  locateParagraph(doc, blockId)?.kind === "cell";

export const textOfRuns = (runs: Run[]): string => runs.map((r) => r.text).join("");

export const textOfBlock = (doc: Document, blockId: string): string => {
  const b = blockById(doc, blockId);
  return b ? textOfRuns(b.runs) : "";
};

export function prevGrapheme(text: string, offset: number): number {
  let prev = 0;
  for (const s of graphemes.segment(text)) {
    if (s.index >= offset) break;
    prev = s.index;
  }
  return prev;
}

export function nextGrapheme(text: string, offset: number): number {
  for (const s of graphemes.segment(text)) {
    const end = s.index + s.segment.length;
    if (end > offset) return end;
  }
  return text.length;
}

export function prevWordStart(text: string, offset: number): number {
  let prev = 0;
  for (const s of words.segment(text)) {
    if (s.index >= offset) break;
    if (s.isWordLike) prev = s.index;
  }
  return prev;
}

export function nextWordEnd(text: string, offset: number): number {
  for (const s of words.segment(text)) {
    const end = s.index + s.segment.length;
    if (s.isWordLike && end > offset) return end;
  }
  return text.length;
}

/** Style of the character before `offset` (Word: typing inherits what precedes
 *  the caret), falling back to the first run's style. */
export function styleAtRuns(runs: Run[], offset: number): CharStyle | undefined {
  let cum = 0;
  for (const r of runs) {
    const end = cum + r.text.length;
    // offset-1 falls inside this run -> its style precedes the caret
    if (offset > cum && offset <= end && r.text.length > 0) return r.style;
    cum = end;
  }
  return runs[0]?.style;
}
