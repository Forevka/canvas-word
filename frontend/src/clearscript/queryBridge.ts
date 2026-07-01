// C#-facing query bridge: pure mappers that project the document query API into
// flat, JSON-serializable DTOs the ClearScript host (WordDocument) marshals into
// C# records. Kept separate from entry.ts so the mapping is unit-testable without
// the V8 bundle. Shapes here MUST stay in lockstep with the C# records in
// dotnet/src/WordCanvas.ClearScript/WordDocumentQuery.cs.

import type { Document, Paragraph } from "@cw/shared";
import { findParagraphs, getSections, textOfRuns, walk, type BlockContext } from "@cw/shared";
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
