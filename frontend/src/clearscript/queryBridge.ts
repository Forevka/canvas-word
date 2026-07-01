// C#-facing query bridge: pure mappers that project the document query API into
// flat, JSON-serializable DTOs the ClearScript host (WordDocument) marshals into
// C# records. Kept separate from entry.ts so the mapping is unit-testable without
// the V8 bundle. Shapes here MUST stay in lockstep with the C# records in
// dotnet/src/WordCanvas.ClearScript/WordDocumentQuery.cs.

import type { Document, DocSelection, Paragraph, Run } from "@cw/shared";
import {
  blockPath,
  findParagraphs,
  getBookmarks,
  getEndnotes,
  getFields,
  getFootnotes,
  getListItems,
  getSdtNodes,
  getSdtValue,
  getSections,
  getStyles,
  positionOfText,
  rangeText,
  textOf,
  textOfRuns,
  walk,
  type BlockContext,
} from "@cw/shared";
import type { PageInfo as LayoutPageInfo } from "../layout/pages";

/** A paragraph flattened for the host: text + where it lives (container / table
 *  cell / note). `row`/`col` are -1 when not in a cell; nullable fields are null
 *  (not undefined) so the C# side reads a stable shape. */
export interface ParagraphInfo {
  id: string;
  text: string;
  container: string;
  tableId: string | null;
  row: number;
  col: number;
  noteKind: string | null;
  noteId: string | null;
  styleName: string | null;
  outlineLevel: number | null;
}

export interface SectionInfo {
  index: number;
  startBlock: number;
  endBlock: number;
  breakType: string;
  pageWidthPx: number;
  pageHeightPx: number;
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
  columnCount: number;
}

export interface PageInfo {
  index: number;
  number: number;
  widthPx: number;
  heightPx: number;
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
  blockIds: string[];
}

/** A content control flattened for the host: its properties, its place in the
 *  nesting tree, and its enclosed text — one call yields the whole SDT forest.
 *  Nullable fields are null (not undefined) so the C# side reads a stable shape. */
export interface SdtInfo {
  id: string;
  sdtType: string;
  tag: string | null;
  alias: string | null;
  /** Checkbox state (null for non-checkbox controls). */
  checked: boolean | null;
  placeholder: boolean;
  /** Parent control id, or null for a top-level control. */
  parentId: string | null;
  /** Direct child control ids, first-appearance order. */
  childIds: string[];
  /** Full ancestry outer→inner ending at this id. */
  path: string[];
  /** Nesting depth (0 = root). */
  depth: number;
  /** The plain text the control encloses. */
  text: string;
}

function paragraphInfo(p: Paragraph, ctx: BlockContext): ParagraphInfo {
  return {
    id: p.id,
    text: textOfRuns(p.runs),
    container: ctx.container,
    tableId: ctx.cell ? ctx.cell.tableId : null,
    row: ctx.cell ? ctx.cell.row : -1,
    col: ctx.cell ? ctx.cell.col : -1,
    noteKind: ctx.note ? ctx.note.kind : null,
    noteId: ctx.note ? ctx.note.id : null,
    styleName: p.style.namedStyle ?? null,
    outlineLevel: p.style.outlineLevel ?? null,
  };
}

/** Every paragraph in the document (body, cells, bands, notes). */
export function queryParagraphs(doc: Document): ParagraphInfo[] {
  const out: ParagraphInfo[] = [];
  walk(doc, (block, ctx) => {
    if (block.kind === "paragraph") out.push(paragraphInfo(block, ctx));
  });
  return out;
}

/** Paragraphs whose text contains `needle` (substring match). */
export function findText(doc: Document, needle: string): ParagraphInfo[] {
  return findParagraphs(doc, needle).map((m) => paragraphInfo(m.paragraph, m.context));
}

export function querySections(doc: Document): SectionInfo[] {
  return getSections(doc).map((s) => ({
    index: s.index,
    startBlock: s.startBlock,
    endBlock: s.endBlock,
    breakType: s.breakType,
    pageWidthPx: s.props.pageWidthPx,
    pageHeightPx: s.props.pageHeightPx,
    marginTop: s.props.marginPx.top,
    marginRight: s.props.marginPx.right,
    marginBottom: s.props.marginPx.bottom,
    marginLeft: s.props.marginPx.left,
    columnCount: s.props.columns?.count ?? 1,
  }));
}

/** Enclosed text for EVERY control in one document walk, keyed by control id —
 *  the batched equivalent of calling `sdtText(doc, id)` per control (which would
 *  re-walk the whole doc each time, O(controls × blocks)). A control is block-level
 *  OR inline, never both, so exactly one branch contributes per id; parts join with
 *  newlines, matching `sdtText`. */
function sdtTextMap(doc: Document): Map<string, string[]> {
  const parts = new Map<string, string[]>();
  const push = (id: string, text: string): void => {
    const a = parts.get(id);
    if (a) a.push(text);
    else parts.set(id, [text]);
  };
  walk(doc, (b) => {
    if (b.sdtPath) {
      const t = textOf(b); // block-level members contribute the whole block
      if (t.length > 0) for (const id of b.sdtPath) push(id, t);
    }
    if (b.kind === "paragraph") {
      // Inline members contribute their runs, grouped per control, per paragraph.
      const perId = new Map<string, Run[]>();
      for (const run of b.runs) {
        const ip = run.style.sdtPath;
        if (!ip) continue;
        for (const id of ip) {
          const arr = perId.get(id);
          if (arr) arr.push(run);
          else perId.set(id, [run]);
        }
      }
      for (const [id, runs] of perId) push(id, textOfRuns(runs));
    }
  });
  return parts;
}

/** Every content control, flattened with its nesting links and enclosed text —
 *  the SDT forest in one call (the primary templating surface, for .NET hosts). */
export function querySdts(doc: Document): SdtInfo[] {
  const textById = sdtTextMap(doc);
  return getSdtNodes(doc).map((n) => ({
    id: n.id,
    sdtType: n.props.type,
    tag: n.props.tag ?? null,
    alias: n.props.alias ?? null,
    checked: n.props.type === "checkbox" ? (n.props.checked ?? false) : null,
    placeholder: n.props.placeholder ?? false,
    parentId: n.parentId ?? null,
    childIds: n.childIds,
    path: n.path,
    depth: n.depth,
    text: (textById.get(n.id) ?? []).join("\n"),
  }));
}

/** Flatten a layout page map (from getPages) into host DTOs. */
export function mapPages(pages: LayoutPageInfo[]): PageInfo[] {
  return pages.map((p) => ({
    index: p.index,
    number: p.number,
    widthPx: p.widthPx,
    heightPx: p.heightPx,
    marginTop: p.marginPx.top,
    marginRight: p.marginPx.right,
    marginBottom: p.marginPx.bottom,
    marginLeft: p.marginPx.left,
    blockIds: p.blockIds,
  }));
}

// --- fields / bookmarks / notes / lists / styles / block-path / text ----------
// Each mapper flattens a `@cw/shared` query getter into a host DTO. Shapes MUST
// stay in lockstep with the C# records in WordDocumentQuery.cs.

export interface FieldInfo {
  id: string;
  name: string;
  kind: string;
  instruction: string;
}

export function queryFields(doc: Document): FieldInfo[] {
  return getFields(doc).map((f) => ({ id: f.id, name: f.name, kind: f.kind, instruction: f.instruction }));
}

export interface BookmarkInfo {
  name: string;
  startBlockId: string;
  startOffset: number;
  endBlockId: string;
  endOffset: number;
}

export function queryBookmarks(doc: Document): BookmarkInfo[] {
  return getBookmarks(doc).map((b) => ({
    name: b.name,
    startBlockId: b.range.start.blockId,
    startOffset: b.range.start.offset,
    endBlockId: b.range.end.blockId,
    endOffset: b.range.end.offset,
  }));
}

/** A footnote/endnote story flattened: its ref id + concatenated text. */
export interface NoteInfo {
  id: string;
  text: string;
}

export function queryFootnotes(doc: Document): NoteInfo[] {
  return getFootnotes(doc).map((n) => ({ id: n.id, text: n.paragraphs.map(textOf).join("\n") }));
}

export function queryEndnotes(doc: Document): NoteInfo[] {
  return getEndnotes(doc).map((n) => ({ id: n.id, text: n.paragraphs.map(textOf).join("\n") }));
}

export interface ListItemInfo {
  paragraphId: string;
  level: number;
  marker: string;
  text: string;
  container: string;
}

export function queryListItems(doc: Document, listId: string): ListItemInfo[] {
  return getListItems(doc, listId).map((li) => ({
    paragraphId: li.paragraph.id,
    level: li.level,
    marker: li.marker,
    text: textOfRuns(li.paragraph.runs),
    container: li.context.container,
  }));
}

export interface StyleInfo {
  id: string;
  name: string;
  styleType: string;
  basedOn: string | null;
}

export function queryStyles(doc: Document): StyleInfo[] {
  return getStyles(doc).map((s) => ({ id: s.id, name: s.name, styleType: s.type ?? "paragraph", basedOn: s.basedOn ?? null }));
}

/** A block's location. `row`/`col` are -1 when not in a cell; nullable fields null. */
export interface BlockPathInfo {
  container: string;
  tableId: string | null;
  row: number;
  col: number;
  noteKind: string | null;
  noteId: string | null;
}

export function queryBlockPath(doc: Document, id: string): BlockPathInfo | null {
  const ctx = blockPath(doc, id);
  if (!ctx) return null;
  return {
    container: ctx.container,
    tableId: ctx.cell ? ctx.cell.tableId : null,
    row: ctx.cell ? ctx.cell.row : -1,
    col: ctx.cell ? ctx.cell.col : -1,
    noteKind: ctx.note ? ctx.note.kind : null,
    noteId: ctx.note ? ctx.note.id : null,
  };
}

export interface PositionInfo {
  blockId: string;
  offset: number;
}

export function queryPositionOfText(doc: Document, needle: string): PositionInfo | null {
  const pos = positionOfText(doc, needle);
  return pos ? { blockId: pos.blockId, offset: pos.offset } : null;
}

/** Concatenated text of a selection, built from two `DocPosition`s. */
export function queryRangeText(doc: Document, startBlockId: string, startOffset: number, endBlockId: string, endOffset: number): string {
  const selection: DocSelection = { anchor: { blockId: startBlockId, offset: startOffset }, focus: { blockId: endBlockId, offset: endOffset } };
  return rangeText(doc, selection);
}

/** A content control's typed value flattened; null if the control does not exist. */
export interface SdtValueInfo {
  sdtType: string;
  text: string;
  checked: boolean | null;
  selected: string | null;
}

export function querySdtValue(doc: Document, id: string): SdtValueInfo | null {
  const v = getSdtValue(doc, id);
  if (!v) return null;
  return { sdtType: v.type, text: v.text, checked: v.checked ?? null, selected: v.selected ?? null };
}
