// Merge / append documents — concatenate a SOURCE document's content after a
// DESTINATION's, in the pure model, reconciling every id space so the two can't
// collide or alias — the headless equivalent of Word's "insert file at end",
// including its style import. See MERGE_PLAN.md.
//
// The model is CONCRETE (formatting baked on runs) but has many keyed registries
// (stylesheet, lists, tableStyles, sdts, fields, footnotes/endnotes, bookmarks)
// and cross-references by id (block ids, sdtPath, fieldId, note refs, list refs,
// style refs, tocEntry targets, bookmark positions). A correct merge appends
// `source.blocks` after `dest.blocks` while REBASING every space: destination ids
// are left untouched (so existing references and caller-held ids stay valid) and
// every SOURCE id is remapped.
//
// Media needs no remapping here: ImageBlock.mediaId is content-addressed
// (sha256), so identical images across documents collapse to one id for free; the
// C# binding unions the two `cw-media` byte maps at the boundary.

import { BAND_CONTAINERS } from "./document";
import type {
  Block,
  CharStyle,
  CustomBlock,
  Document,
  EquationBlock,
  FieldDef,
  ImageBlock,
  Paragraph,
  ParaStyle,
  Run,
  SectionBreakType,
  SectionPatch,
  SectionProps,
  TableBlock,
  TableCell,
  TableRow,
} from "./document";
import { DEFAULT_CHAR_STYLE, DEFAULT_PARA_STYLE } from "./defaults";
import type { DocPosition } from "./position";
import type { NamedStyle, Stylesheet } from "./stylesheet";
import type { TableStyle } from "./tableStyles";
import { normSdtPath } from "./sdt";
import { freshId } from "../ids";

/** How to reconcile named styles / table styles that collide with the destination.
 *  "useDestination" (default, ~ Word's UseDestinationStyles): a source style whose
 *  NAME matches a destination style is dropped and its references repointed at the
 *  destination's; source-only styles are added (renaming the id if the id — not the
 *  name — collides). Run formatting is untouched either way (the model is concrete),
 *  so text keeps its look; this only picks which style ENTITY the round-trip
 *  reference points at. "keepSource": always keep the source's own styles, renaming
 *  ids on collision, so the source's definitions win for its content. */
export type StyleMergeMode = "useDestination" | "keepSource";

/** The section seam between the destination's last section and the source's first.
 *  "nextPage"/"evenPage"/"oddPage": insert a section break so the source starts on a
 *  new (parity-forced) page and each side keeps its OWN section geometry + bands.
 *  "continuous"/"none": append inline under the destination's section (the source's
 *  final section geometry/bands are dropped — mid-document breaks inside the source
 *  are always preserved). Default "nextPage". */
export type MergeSectionBreak = "nextPage" | "evenPage" | "oddPage" | "continuous" | "none";

export interface MergeOptions {
  styles?: StyleMergeMode;
  sectionBreak?: MergeSectionBreak;
  /** Rename a source bookmark whose name already exists in the destination instead
   *  of dropping it. Default true. */
  renameBookmarksOnCollision?: boolean;
}

/** old→new id per space, for every SOURCE id that changed or was repointed. */
export interface MergeIdMap {
  blocks: Record<string, string>;
  styles: Record<string, string>;
  lists: Record<string, string>;
  tableStyles: Record<string, string>;
  sdts: Record<string, string>;
  fields: Record<string, string>;
  footnotes: Record<string, string>;
  endnotes: Record<string, string>;
  bookmarks: Record<string, string>;
}

export interface MergeWarning {
  code: string;
  detail?: string;
}

export interface MergeResult {
  /** A NEW document; `dest` and `source` are not mutated. */
  doc: Document;
  idMap: MergeIdMap;
  warnings: MergeWarning[];
}

// ---------------------------------------------------------------------------

/** A unique key derived from `base`, suffixed `__2`, `__3`, … until free. */
function uniqueKey(base: string, taken: ReadonlySet<string>): string {
  if (!taken.has(base)) return base;
  let n = 2;
  while (taken.has(`${base}__${n}`)) n++;
  return `${base}__${n}`;
}

const remap = (map: ReadonlyMap<string, string>, id: string): string => map.get(id) ?? id;

/** Internal remap context: the fully-populated id maps used while cloning content. */
interface Ctx {
  blocks: ReadonlyMap<string, string>;
  styles: ReadonlyMap<string, string>;
  lists: ReadonlyMap<string, string>;
  tableStyles: ReadonlyMap<string, string>;
  sdts: ReadonlyMap<string, string>;
  fields: ReadonlyMap<string, string>;
  footnotes: ReadonlyMap<string, string>;
  endnotes: ReadonlyMap<string, string>;
}

// ---- id collection (pass A) ------------------------------------------------

/** Assign a fresh block id to every block reachable from `blocks` — recursing into
 *  table cells and section-break band stories — accumulating into `into`. */
function collectBlockIds(blocks: readonly Block[], into: Map<string, string>): void {
  for (const b of blocks) {
    into.set(b.id, freshId());
    if (b.kind === "table") {
      for (const row of b.rows) for (const cell of row.cells) collectBlockIds(cell.blocks, into);
    } else if (b.kind === "paragraph" && b.style.sectionBreak) {
      for (const key of BAND_CONTAINERS) {
        const band = b.style.sectionBreak.props[key];
        if (band) collectBlockIds(band, into);
      }
    }
  }
}

// ---- content cloning (pass B) ----------------------------------------------

function cloneChar(style: CharStyle, ctx: Ctx): CharStyle {
  const s: CharStyle = { ...style };
  if (s.sdtPath) s.sdtPath = normSdtPath(s.sdtPath.map((id) => remap(ctx.sdts, id)));
  if (s.fieldId) s.fieldId = remap(ctx.fields, s.fieldId);
  if (s.footnoteRef) s.footnoteRef = remap(ctx.footnotes, s.footnoteRef);
  if (s.endnoteRef) s.endnoteRef = remap(ctx.endnotes, s.endnoteRef);
  if (s.charStyleId) s.charStyleId = remap(ctx.styles, s.charStyleId);
  return s;
}

function cloneRun(run: Run, ctx: Ctx): Run {
  return { text: run.text, style: cloneChar(run.style, ctx) };
}

function cloneParaStyle(style: ParaStyle, ctx: Ctx): ParaStyle {
  const s: ParaStyle = { ...style };
  if (s.namedStyle) s.namedStyle = remap(ctx.styles, s.namedStyle);
  if (s.list) s.list = { ...s.list, listId: remap(ctx.lists, s.list.listId) };
  if (s.tocEntry) s.tocEntry = { ...s.tocEntry, targetId: remap(ctx.blocks, s.tocEntry.targetId) };
  if (s.sectionBreak) {
    s.sectionBreak = { type: s.sectionBreak.type, props: cloneSectionPatch(s.sectionBreak.props, ctx) };
  }
  return s;
}

function cloneSectionPatch(patch: SectionPatch, ctx: Ctx): SectionPatch {
  const p: SectionPatch = { ...patch };
  for (const key of BAND_CONTAINERS) {
    const band = p[key];
    if (band) p[key] = band.map((b) => cloneBlock(b, ctx));
  }
  return p;
}

/** Remap an sdt/field path (each id via the given map); undefined when empty. */
function clonePath(path: string[] | undefined, map: ReadonlyMap<string, string>): string[] | undefined {
  return path ? normSdtPath(path.map((id) => remap(map, id))) : undefined;
}

function cloneCell(cell: TableCell, ctx: Ctx): TableCell {
  return { ...cell, id: freshId(), blocks: cell.blocks.map((b) => cloneBlock(b, ctx)) };
}

function cloneRow(row: TableRow, ctx: Ctx): TableRow {
  return { ...row, cells: row.cells.map((c) => cloneCell(c, ctx)) };
}

function cloneBlock(block: Block, ctx: Ctx): Block {
  const id = remap(ctx.blocks, block.id);
  switch (block.kind) {
    case "paragraph": {
      const out: Paragraph = { ...block, id, runs: block.runs.map((r) => cloneRun(r, ctx)), style: cloneParaStyle(block.style, ctx) };
      if (out.fieldId) out.fieldId = remap(ctx.fields, out.fieldId);
      out.sdtPath = clonePath(block.sdtPath, ctx.sdts);
      return out;
    }
    case "image": {
      const out: ImageBlock = { ...block, id }; // src + mediaId (content-addressed) unchanged
      if (out.fieldId) out.fieldId = remap(ctx.fields, out.fieldId);
      out.sdtPath = clonePath(block.sdtPath, ctx.sdts);
      return out;
    }
    case "table": {
      const out: TableBlock = { ...block, id, rows: block.rows.map((r) => cloneRow(r, ctx)) };
      if (out.styleId) out.styleId = remap(ctx.tableStyles, out.styleId);
      if (out.fieldId) out.fieldId = remap(ctx.fields, out.fieldId);
      out.sdtPath = clonePath(block.sdtPath, ctx.sdts);
      return out;
    }
    case "equation": {
      const out: EquationBlock = { ...block, id }; // equation AST is immutable — shared
      if (out.fieldId) out.fieldId = remap(ctx.fields, out.fieldId);
      out.sdtPath = clonePath(block.sdtPath, ctx.sdts);
      return out;
    }
    case "custom": {
      // data is JSON-serializable by contract — deep-clone it so the copy never
      // aliases the original block's payload.
      const out: CustomBlock = { ...block, id, data: block.data === undefined ? undefined : structuredClone(block.data) };
      if (out.fieldId) out.fieldId = remap(ctx.fields, out.fieldId);
      out.sdtPath = clonePath(block.sdtPath, ctx.sdts);
      return out;
    }
  }
}

function cloneField(def: FieldDef, newId: string, ctx: Ctx): FieldDef {
  const out: FieldDef = { ...def, id: newId };
  if (out.spec?.type === "IF") {
    out.spec = {
      ...out.spec,
      trueRuns: out.spec.trueRuns.map((r) => cloneRun(r, ctx)),
      falseRuns: out.spec.falseRuns.map((r) => cloneRun(r, ctx)),
    };
  }
  return out;
}

function cloneSection(section: SectionProps, ctx: Ctx): SectionProps {
  const out: SectionProps = { ...section };
  for (const key of BAND_CONTAINERS) {
    const band = out[key];
    if (band) out[key] = band.map((b) => cloneBlock(b, ctx));
  }
  return out;
}

// ---- section seam ----------------------------------------------------------

/** A SectionPatch that FULLY describes `props`, so a block governed by it never
 *  inherits the merged document's (source's) section. Every inheritable field —
 *  columns and all six bands included — is set explicitly (empty band ⇒ []). */
function sectionPropsToPatch(props: SectionProps): SectionPatch {
  const patch: SectionPatch = {
    pageWidthPx: props.pageWidthPx,
    pageHeightPx: props.pageHeightPx,
    marginPx: props.marginPx,
    columns: props.columns ?? null,
  };
  if (props.pageNumberStart !== undefined) patch.pageNumberStart = props.pageNumberStart;
  if (props.pageColorHex !== undefined) patch.pageColorHex = props.pageColorHex;
  if (props.pageBorders !== undefined) patch.pageBorders = props.pageBorders;
  if (props.headerDistancePx !== undefined) patch.headerDistancePx = props.headerDistancePx;
  if (props.footerDistancePx !== undefined) patch.footerDistancePx = props.footerDistancePx;
  if (props.lineNumbering !== undefined) patch.lineNumbering = props.lineNumbering;
  for (const key of BAND_CONTAINERS) patch[key] = props[key] ?? [];
  return patch;
}

/** An empty paragraph that TERMINATES the destination's section (carrying its full
 *  geometry + bands), starting the source's first section on a new page. */
function seamParagraph(destSection: SectionProps, type: SectionBreakType): Paragraph {
  return {
    kind: "paragraph",
    id: freshId(),
    revision: 0,
    runs: [{ text: "", style: { ...DEFAULT_CHAR_STYLE } }],
    style: { ...DEFAULT_PARA_STYLE, sectionBreak: { type, props: sectionPropsToPatch(destSection) } },
  };
}

// ---- registry reconciliation (pass A) --------------------------------------

/** A named, id-keyed, optionally basedOn-chained registry entry (NamedStyle / TableStyle). */
interface NamedEntity {
  id: string;
  name: string;
  basedOn?: string;
}

/** Reconcile named, id-keyed entities from `source` into `dest`: in "useDestination"
 *  a same-NAMED source entity is dropped and repointed at the destination's; otherwise
 *  it's added, renaming its id on collision. Populates `map` (source id → resulting id)
 *  and remaps each added entity's `basedOn`. Returns the merged list (destination
 *  entities first, then the added source ones). Shared by styles + table styles. */
function mergeNamedEntities<T extends NamedEntity>(
  dest: readonly T[],
  source: readonly T[],
  mode: StyleMergeMode,
  map: Map<string, string>,
  warnings: MergeWarning[],
  warnCode: string,
): T[] {
  const out: T[] = dest.map((e) => ({ ...e }));
  const taken = new Set(out.map((e) => e.id));
  const destIdByName = new Map(out.map((e) => [e.name, e.id] as const));
  const added: T[] = [];
  for (const e of source) {
    if (mode === "useDestination") {
      const destId = destIdByName.get(e.name);
      if (destId !== undefined) {
        map.set(e.id, destId);
        continue;
      }
    }
    let newId = e.id;
    if (taken.has(newId)) {
      newId = uniqueKey(e.id, taken);
      warnings.push({ code: warnCode, detail: `${e.id} → ${newId}` });
    }
    taken.add(newId);
    map.set(e.id, newId);
    added.push({ ...e, id: newId });
  }
  for (const e of added) if (e.basedOn !== undefined) e.basedOn = remap(map, e.basedOn);
  out.push(...added);
  return out;
}

function mergeStylesheets(
  dest: Stylesheet | undefined,
  source: Stylesheet | undefined,
  mode: StyleMergeMode,
  styleIdMap: Map<string, string>,
  warnings: MergeWarning[],
): Stylesheet | undefined {
  if (!source) return dest;
  const styles = mergeNamedEntities<NamedStyle>(dest?.styles ?? [], source.styles, mode, styleIdMap, warnings, "style-id-renamed");
  return { styles, defaultStyleId: dest?.defaultStyleId ?? source.defaultStyleId };
}

function mergeTableStyles(
  dest: Record<string, TableStyle> | undefined,
  source: Record<string, TableStyle> | undefined,
  mode: StyleMergeMode,
  map: Map<string, string>,
  warnings: MergeWarning[],
): Record<string, TableStyle> {
  const merged = mergeNamedEntities<TableStyle>(Object.values(dest ?? {}), Object.values(source ?? {}), mode, map, warnings, "table-style-id-renamed");
  const out: Record<string, TableStyle> = {};
  for (const t of merged) out[t.id] = t;
  return out;
}

/** Rename-on-collision merge for a keyed record (lists / sdts), remapping every
 *  source key to a fresh one when it collides with the destination. */
function mergeKeyed<T>(
  dest: Record<string, T> | undefined,
  source: Record<string, T> | undefined,
  map: Map<string, string>,
): Record<string, T> {
  const out: Record<string, T> = { ...(dest ?? {}) };
  const taken = new Set(Object.keys(out));
  for (const key of Object.keys(source ?? {})) {
    let newKey = key;
    if (taken.has(newKey)) newKey = uniqueKey(key, taken);
    taken.add(newKey);
    map.set(key, newKey);
  }
  return out;
}

// ---------------------------------------------------------------------------

function emptyIdMap(): MergeIdMap {
  return { blocks: {}, styles: {}, lists: {}, tableStyles: {}, sdts: {}, fields: {}, footnotes: {}, endnotes: {}, bookmarks: {} };
}

/** The result of reconciling `source`'s registries + id spaces INTO `dest`: the
 *  merged registry records, the source's body blocks + section cloned through the
 *  new id maps, and the id map / warnings. Shared by `mergeDocuments` (append) and
 *  `replaceSdtContent` (splice into a control). Does NOT decide where the blocks go. */
export interface Reconciled {
  stylesheet: Document["stylesheet"];
  lists: NonNullable<Document["lists"]>;
  tableStyles: NonNullable<Document["tableStyles"]>;
  sdts: NonNullable<Document["sdts"]>;
  fields: NonNullable<Document["fields"]>;
  footnotes: NonNullable<Document["footnotes"]>;
  endnotes: NonNullable<Document["endnotes"]>;
  bookmarks: NonNullable<Document["bookmarks"]>;
  /** The source's body blocks, cloned through the new id maps (ready to place). */
  sourceBlocks: Block[];
  /** The source's section, cloned through the new id maps (used when it governs a
   *  merged tail section; ignored when the source flows into the destination's). */
  sourceSection: SectionProps;
  idMap: MergeIdMap;
  warnings: MergeWarning[];
}

/** Reconcile `source` into `dest`: allocate + remap every id space, merge the
 *  registries, and clone the source's blocks/section/notes/bookmarks through the
 *  resulting maps. Pure — reads both docs, mutates neither. */
export function reconcileSource(dest: Document, source: Document, opts: MergeOptions = {}): Reconciled {
  const mode: StyleMergeMode = opts.styles ?? "useDestination";
  const renameBookmarks = opts.renameBookmarksOnCollision ?? true;
  const warnings: MergeWarning[] = [];

  // ---- Pass A: allocate every id space (keys only; content cloned in pass B) ----
  const blockIdMap = new Map<string, string>();
  collectBlockIds(source.blocks, blockIdMap);
  for (const key of BAND_CONTAINERS) {
    const band = source.section[key];
    if (band) collectBlockIds(band, blockIdMap);
  }
  for (const paras of Object.values(source.footnotes ?? {})) collectBlockIds(paras, blockIdMap);
  for (const paras of Object.values(source.endnotes ?? {})) collectBlockIds(paras, blockIdMap);

  const styleIdMap = new Map<string, string>();
  const stylesheet = mergeStylesheets(dest.stylesheet, source.stylesheet, mode, styleIdMap, warnings);

  const listIdMap = new Map<string, string>();
  const mergedLists = mergeKeyed(dest.lists, source.lists, listIdMap);
  for (const [oldId, newId] of listIdMap) mergedLists[newId] = { ...source.lists![oldId]!, id: newId };

  const tableStyleIdMap = new Map<string, string>();
  const mergedTableStyles = mergeTableStyles(dest.tableStyles, source.tableStyles, mode, tableStyleIdMap, warnings);

  const sdtIdMap = new Map<string, string>();
  const mergedSdts = mergeKeyed(dest.sdts, source.sdts, sdtIdMap);
  for (const [oldId, newId] of sdtIdMap) mergedSdts[newId] = { ...source.sdts![oldId]! };

  const fieldIdMap = new Map<string, string>();
  mergeKeyed(dest.fields, source.fields, fieldIdMap); // allocate keys only; bodies cloned below

  const footnoteIdMap = new Map<string, string>();
  const mergedFootnotes: Record<string, Paragraph[]> = { ...(dest.footnotes ?? {}) };
  mergeKeyed(dest.footnotes, source.footnotes, footnoteIdMap);

  const endnoteIdMap = new Map<string, string>();
  const mergedEndnotes: Record<string, Paragraph[]> = { ...(dest.endnotes ?? {}) };
  mergeKeyed(dest.endnotes, source.endnotes, endnoteIdMap);

  const ctx: Ctx = {
    blocks: blockIdMap, styles: styleIdMap, lists: listIdMap, tableStyles: tableStyleIdMap,
    sdts: sdtIdMap, fields: fieldIdMap, footnotes: footnoteIdMap, endnotes: endnoteIdMap,
  };

  // ---- Pass B: clone content through the now-complete maps ----
  const mergedFields: Record<string, FieldDef> = { ...(dest.fields ?? {}) };
  for (const [oldId, newId] of fieldIdMap) mergedFields[newId] = cloneField(source.fields![oldId]!, newId, ctx);

  for (const [oldId, newId] of footnoteIdMap) mergedFootnotes[newId] = source.footnotes![oldId]!.map((p) => cloneBlock(p, ctx) as Paragraph);
  for (const [oldId, newId] of endnoteIdMap) mergedEndnotes[newId] = source.endnotes![oldId]!.map((p) => cloneBlock(p, ctx) as Paragraph);

  const remapPos = (pos: DocPosition): DocPosition => ({ blockId: remap(blockIdMap, pos.blockId), offset: pos.offset });
  const bookmarkNameMap = new Map<string, string>();
  const mergedBookmarks: NonNullable<Document["bookmarks"]> = { ...(dest.bookmarks ?? {}) };
  const takenBookmarks = new Set(Object.keys(mergedBookmarks));
  for (const [name, range] of Object.entries(source.bookmarks ?? {})) {
    let newName = name;
    if (takenBookmarks.has(newName)) {
      if (!renameBookmarks) {
        warnings.push({ code: "bookmark-dropped", detail: name });
        continue;
      }
      newName = uniqueKey(name, takenBookmarks);
      warnings.push({ code: "bookmark-renamed", detail: `${name} → ${newName}` });
    }
    takenBookmarks.add(newName);
    bookmarkNameMap.set(name, newName);
    mergedBookmarks[newName] = { start: remapPos(range.start), end: remapPos(range.end) };
  }

  const asRecord = (m: ReadonlyMap<string, string>): Record<string, string> => Object.fromEntries(m);
  return {
    stylesheet,
    lists: mergedLists,
    tableStyles: mergedTableStyles,
    sdts: mergedSdts,
    fields: mergedFields,
    footnotes: mergedFootnotes,
    endnotes: mergedEndnotes,
    bookmarks: mergedBookmarks,
    sourceBlocks: source.blocks.map((b) => cloneBlock(b, ctx)),
    sourceSection: cloneSection(source.section, ctx),
    idMap: {
      blocks: asRecord(blockIdMap),
      styles: asRecord(styleIdMap),
      lists: asRecord(listIdMap),
      tableStyles: asRecord(tableStyleIdMap),
      sdts: asRecord(sdtIdMap),
      fields: asRecord(fieldIdMap),
      footnotes: asRecord(footnoteIdMap),
      endnotes: asRecord(endnoteIdMap),
      bookmarks: asRecord(bookmarkNameMap),
    },
    warnings,
  };
}

/** Build a new document from `dest` with the reconciled registries folded in and an
 *  explicit `section` + `blocks`. Document-level scalars follow the destination. */
export function foldRegistries(dest: Document, r: Reconciled, section: SectionProps, blocks: Block[]): Document {
  const doc: Document = { section, blocks };
  if (r.stylesheet) doc.stylesheet = r.stylesheet;
  if (Object.keys(r.lists).length) doc.lists = r.lists;
  if (Object.keys(r.tableStyles).length) doc.tableStyles = r.tableStyles;
  if (Object.keys(r.footnotes).length) doc.footnotes = r.footnotes;
  if (Object.keys(r.endnotes).length) doc.endnotes = r.endnotes;
  if (Object.keys(r.sdts).length) doc.sdts = r.sdts;
  if (Object.keys(r.bookmarks).length) doc.bookmarks = r.bookmarks;
  if (Object.keys(r.fields).length) doc.fields = r.fields;
  // Document-level scalars: destination wins.
  if (dest.tocInstruction !== undefined) doc.tocInstruction = dest.tocInstruction;
  if (dest.tocAnchorBlockId !== undefined) doc.tocAnchorBlockId = dest.tocAnchorBlockId;
  if (dest.defaultTabStopPx !== undefined) doc.defaultTabStopPx = dest.defaultTabStopPx;
  if (dest.compatSettings) doc.compatSettings = dest.compatSettings;
  return doc;
}

/** Append `source` after `dest`, reconciling every id space. Pure — returns a new
 *  document; the inputs are not mutated. */
export function mergeDocuments(dest: Document, source: Document, opts: MergeOptions = {}): MergeResult {
  const seam: MergeSectionBreak = opts.sectionBreak ?? "nextPage";

  // Nothing to append — return the destination as a fresh top-level object.
  if (source.blocks.length === 0) {
    return { doc: { ...dest }, idMap: emptyIdMap(), warnings: [{ code: "empty-source" }] };
  }

  const r = reconcileSource(dest, source, opts);

  // ---- seam + block assembly ----
  let blocks: Block[];
  let section: SectionProps;
  if (dest.blocks.length === 0) {
    blocks = r.sourceBlocks;
    section = r.sourceSection; // adopt the source's section wholesale
  } else if (seam === "continuous" || seam === "none") {
    blocks = [...dest.blocks, ...r.sourceBlocks];
    section = dest.section; // destination governs the trailing (now-extended) section
    r.warnings.push({ code: "source-section-dropped", detail: seam });
  } else {
    blocks = [...dest.blocks, seamParagraph(dest.section, seam), ...r.sourceBlocks];
    section = r.sourceSection; // source governs its own tail section
  }

  return { doc: foldRegistries(dest, r, section, blocks), idMap: r.idMap, warnings: r.warnings };
}

/** Fold N documents left-to-right: `mergeAll([a, b, c])` ≡ merge(merge(a, b), c).
 *  Returns the FINAL merge's id map (the last source's remap) and the accumulated
 *  warnings. Throws on an empty list. */
export function mergeAll(docs: Document[], opts: MergeOptions = {}): MergeResult {
  if (docs.length === 0) throw new Error("mergeAll requires at least one document");
  let doc = docs[0]!;
  let idMap = emptyIdMap();
  const warnings: MergeWarning[] = [];
  for (let i = 1; i < docs.length; i++) {
    const r = mergeDocuments(doc, docs[i]!, opts);
    doc = r.doc;
    idMap = r.idMap;
    warnings.push(...r.warnings);
  }
  return { doc, idMap, warnings };
}
