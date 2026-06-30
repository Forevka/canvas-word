// Shared text helpers over the model: paragraph text access, grapheme/word
// boundaries (Intl.Segmenter — the same segmentation pretext breaks lines on),
// and style-at-position for Word-style inheritance when typing.

import type { BandContainer, Block, CharStyle, Document, Paragraph, Run } from "./document";
import { BAND_CONTAINERS } from "./document";

export const graphemes = new Intl.Segmenter(undefined, { granularity: "grapheme" });
export const words = new Intl.Segmenter(undefined, { granularity: "word" });

/** Step-0 perf diagnostic (off by default, ~free when disabled). When `enabled`,
 *  paragraphsOf/bandParagraphs tally how often they rebuild the paragraph list and
 *  how many paragraph slots they walk+allocate — the exact work a WeakMap index
 *  keyed on document identity would eliminate. Toggle via the export, read the
 *  counters, reset between workloads. Used by frontend/src/perf/readsPerDoc.test.ts. */
export const _paraScanStats = {
  enabled: false,
  /** # of full-document traversals (paragraphsOf calls). */
  fullCalls: 0,
  /** Sum of full-list lengths returned (body + band paragraphs walked). */
  fullSlots: 0,
  /** # of bandParagraphs calls (incl. the O(N²) isBandParagraph filter). */
  bandCalls: 0,
  /** Sum of band-list lengths walked. */
  bandSlots: 0,
};
export const resetParaScanStats = (): void => {
  _paraScanStats.fullCalls = 0;
  _paraScanStats.fullSlots = 0;
  _paraScanStats.bandCalls = 0;
  _paraScanStats.bandSlots = 0;
};

/** A paragraph whose every non-empty run is hidden (w:vanish): it has text but
 *  none is visible. Such a paragraph is never laid out, caret-reachable, or
 *  deletable — it's preserved metadata. A truly empty paragraph (no text at all)
 *  is NOT hidden — it's a normal blank line. */
export const isHiddenParagraph = (p: Paragraph): boolean =>
  p.runs.some((r) => r.text.length > 0) && p.runs.every((r) => r.text.length === 0 || r.style.hidden === true);

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
export const bandParagraphs = (doc: Document, band: BandContainer): Paragraph[] => {
  const list = paragraphsInBlocks(doc.section[band] ?? [], true);
  if (_paraScanStats.enabled) {
    _paraScanStats.bandCalls++;
    _paraScanStats.bandSlots += list.length;
  }
  return list;
};

/** Body paragraphs in document order, INCLUDING table-cell paragraphs (one
 *  level deep). Excludes margin bands and footnote notes — the body story only.
 *  Heading/TOC scans use this so headings inside table cells are found too. */
export const bodyParagraphs = (doc: Document): Paragraph[] => paragraphsInBlocks(doc.blocks, true);

/** Per-document memoized index of the navigable-paragraph space. The model is
 *  immutable — every edit returns a NEW Document object (path-clone) — so a
 *  WeakMap keyed on document identity auto-invalidates: a fresh doc simply misses
 *  the cache and the old entry is GC'd. Built once per doc, then `paragraphsOf`
 *  is a cached-array return and `blockById`/`blockIndexOf` are O(1) Map lookups,
 *  instead of re-walking the whole block tree (body + table cells + 6 bands +
 *  footnotes/endnotes) on every call. Mirrors layout/geometry.ts's TreeIndex. */
interface DocParaIndex {
  list: Paragraph[];
  /** id → its index in `list` (FIRST occurrence, matching the old `.find`/
   *  `.findIndex` semantics for the rare duplicate id). */
  byId: Map<string, number>;
}
const paraIndexCache = new WeakMap<Document, DocParaIndex>();

const buildParaIndex = (doc: Document): DocParaIndex => {
  const list = [
    ...paragraphsInBlocks(doc.blocks, true),
    ...BAND_CONTAINERS.flatMap((band) => bandParagraphs(doc, band)),
    ...Object.values(doc.footnotes ?? {}).flat(),
    ...Object.values(doc.endnotes ?? {}).flat(),
  ];
  const byId = new Map<string, number>();
  for (let i = 0; i < list.length; i++) {
    const id = list[i]!.id;
    if (!byId.has(id)) byId.set(id, i); // keep first — matches old .find/.findIndex
  }
  if (_paraScanStats.enabled) {
    _paraScanStats.fullCalls++;
    _paraScanStats.fullSlots += list.length;
  }
  return { list, byId };
};

const paraIndex = (doc: Document): DocParaIndex => {
  let idx = paraIndexCache.get(doc);
  if (!idx) {
    idx = buildParaIndex(doc);
    paraIndexCache.set(doc, idx);
  }
  return idx;
};

/** All editable paragraphs in document order: body (including table cells),
 *  then every band story (header/footer + first/even variants). This is the
 *  index space commands use. Cached per document identity (see DocParaIndex). */
export const paragraphsOf = (doc: Document): Paragraph[] => paraIndex(doc).list;

export const blockById = (doc: Document, blockId: string): Paragraph | undefined => {
  const idx = paraIndex(doc);
  const i = idx.byId.get(blockId);
  return i === undefined ? undefined : idx.list[i];
};

export const blockIndexOf = (doc: Document, blockId: string): number => {
  const i = paraIndex(doc).byId.get(blockId);
  return i === undefined ? -1 : i;
};

/** The set of every band (header/footer + variants) paragraph id, memoized per
 *  document identity. Lets a body-vs-band membership test be an O(1) Set lookup
 *  instead of re-walking all six band stories per query — the selection
 *  controller's caret navigation filters the body paragraphs against this on
 *  every arrow key. Same WeakMap-on-identity invalidation as the para index. */
const bandIdCache = new WeakMap<Document, ReadonlySet<string>>();
export const bandParagraphIds = (doc: Document): ReadonlySet<string> => {
  let s = bandIdCache.get(doc);
  if (!s) {
    const built = new Set<string>();
    for (const band of BAND_CONTAINERS) for (const p of bandParagraphs(doc, band)) built.add(p.id);
    s = built;
    bandIdCache.set(doc, s);
  }
  return s;
};

/** Body navigation order: the paragraphs the caret can reach while editing the
 *  body — `paragraphsOf` minus band paragraphs (they belong to their own header/
 *  footer story) and hidden (w:vanish) paragraphs (never laid out, so the caret
 *  skips them). Memoized per document identity: the selection controller derives
 *  this on every caret move, and the filter (a full pass plus a per-paragraph
 *  hidden test) is otherwise re-run several times per keystroke. Same
 *  WeakMap-on-identity invalidation as `paragraphsOf`. Treat the result as
 *  IMMUTABLE (callers read-only; see paragraphsOf). */
const navCache = new WeakMap<Document, Paragraph[]>();
export const navigableParagraphs = (doc: Document): Paragraph[] => {
  let list = navCache.get(doc);
  if (!list) {
    const bandIds = bandParagraphIds(doc);
    list = paragraphsOf(doc).filter((p) => !bandIds.has(p.id) && !isHiddenParagraph(p));
    navCache.set(doc, list);
  }
  return list;
};

/** Story (header/footer editing) navigation order for ONE band container:
 *  `bandParagraphs(doc, band)` minus hidden (w:vanish) paragraphs the caret skips.
 *  The body counterpart is `navigableParagraphs`; this is the per-band analogue the
 *  selection controller derives on every caret move while editing a margin band.
 *  Memoized per (document identity, band) — a nested `Map<BandContainer, …>` per
 *  doc, with the same WeakMap-on-identity invalidation. Treat the result as
 *  IMMUTABLE (callers read-only). */
const navBandCache = new WeakMap<Document, Map<BandContainer, Paragraph[]>>();
export const navigableBandParagraphs = (doc: Document, band: BandContainer): Paragraph[] => {
  let perDoc = navBandCache.get(doc);
  if (!perDoc) {
    perDoc = new Map();
    navBandCache.set(doc, perDoc);
  }
  let list = perDoc.get(band);
  if (!list) {
    list = bandParagraphs(doc, band).filter((p) => !isHiddenParagraph(p));
    perDoc.set(band, list);
  }
  return list;
};

/** Every block id present anywhere in the doc — top-level blocks (paragraph,
 *  image, table) in body + bands, table ids, and paragraph/image ids nested in
 *  table cells. Used by the review layer's structural-suggestion GC, which keys
 *  on whole-block identity rather than a text offset range. */
const collectBlockIds = (blocks: Block[], into: Set<string>): void => {
  for (const b of blocks) {
    into.add(b.id);
    if (b.kind === "table") {
      for (const row of b.rows) for (const cell of row.cells) collectBlockIds(cell.blocks, into);
    }
  }
};

/** Does a block with this id exist anywhere in the document (any kind, any
 *  container, including table cells)? */
export const blockExists = (doc: Document, blockId: string): boolean => {
  const ids = new Set<string>();
  collectBlockIds(doc.blocks, ids);
  for (const band of BAND_CONTAINERS) collectBlockIds(doc.section[band] ?? [], ids);
  for (const paras of Object.values(doc.footnotes ?? {})) for (const p of paras) ids.add(p.id);
  for (const paras of Object.values(doc.endnotes ?? {})) for (const p of paras) ids.add(p.id);
  return ids.has(blockId);
};

// ---------------------------------------------------------------------------
// Paragraph locator — content ops work on any paragraph, wherever it lives.

export type ParaLocation =
  | { kind: "top"; bi: number }
  | { kind: "cell"; where: "body" | BandContainer; bi: number; ri: number; ci: number; pi: number }
  | { kind: "band"; band: BandContainer; bi: number }
  | { kind: "footnote"; noteId: string; pi: number }
  | { kind: "endnote"; noteId: string; pi: number };

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
  for (const [noteId, paras] of Object.entries(doc.endnotes ?? {})) {
    const pi = paras.findIndex((p) => p.id === blockId);
    if (pi >= 0) return { kind: "endnote", noteId, pi };
  }
  return null;
}

export function paragraphAt(doc: Document, loc: ParaLocation): Paragraph {
  if (loc.kind === "top") return doc.blocks[loc.bi] as Paragraph;
  if (loc.kind === "band") return doc.section[loc.band]![loc.bi] as Paragraph;
  if (loc.kind === "footnote") return doc.footnotes![loc.noteId]![loc.pi]!;
  if (loc.kind === "endnote") return doc.endnotes![loc.noteId]![loc.pi]!;
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
  if (loc.kind === "endnote") {
    const paras = doc.endnotes![loc.noteId]!.slice();
    paras[loc.pi] = p;
    return { ...doc, endnotes: { ...doc.endnotes, [loc.noteId]: paras } };
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

/** Style of the character AT `index` (the run covering it), or undefined when out
 *  of range. Distinct from styleAtRuns (which returns the char BEFORE a caret
 *  offset): use this to inspect a specific character, e.g. an inline-equation
 *  sentinel located by its run index. */
export function styleOfCharAt(runs: Run[], index: number): CharStyle | undefined {
  if (index < 0) return undefined;
  let cum = 0;
  for (const r of runs) {
    if (index >= cum && index < cum + r.text.length) return r.style;
    cum += r.text.length;
  }
  return undefined;
}
