// DocumentEditor — an ergonomic, headless editing facade over the `applyOp`
// operation engine, the rough analog of mutating a .NET `WordprocessingDocument`
// and saving. It holds a mutable `doc`, translates high-level calls into the
// existing typed `Op`s, threads them through `applyOp` (which does structural
// sharing and produces exact inverses), and keeps an undo/redo stack for free.
//
// This is deliberately NOT the interactive editor's Command/Transaction/caret
// machinery (that lives in the frontend and is UI-coupled). It is a plain data
// facade usable from Node, the browser, and — later — the C# bindings.

import type { Block, CharStyle, Document, Paragraph, ParaStyle, Run, SdtProps, TableCell, TableRow } from "./document";
import { applyOp, applyStylePatchToRuns, containerOf, gridColumnCount, type Op } from "./ops";
import { findParagraphs, getBlockById, getParagraphById, getSdt, getSdtBlocks, getTableById, textOfCell, walk, type ParagraphMatch } from "./query";
import { ancestryThrough, removeSdt as stripSdtFromPath } from "./sdt";
import { mergeDocuments, type MergeOptions, type MergeResult } from "./merge";
import { resolveStyle } from "./stylesheet";
import { DEFAULT_CHAR_STYLE, DEFAULT_PARA_STYLE } from "./defaults";
import { freshId } from "../ids";

/** One undoable step: the forward ops and their inverses (already in
 *  reverse-apply order, ready to replay for undo). */
interface Commit {
  forward: Op[];
  inverse: Op[];
}

/** Structural/positional paragraph-style markers that must NOT be inherited by a
 *  freshly inserted paragraph (they'd conjure a spurious section/page/column
 *  break or a stale TOC entry). */
const NON_INHERITED_STYLE_KEYS = ["sectionBreak", "pageBreakBefore", "columnBreakBefore", "tocEntry"] as const;

export interface InsertParagraphOptions {
  /** Insert before or after the reference block (default "after"). */
  position?: "before" | "after";
  /** Paragraph style for the new block. Defaults to the reference paragraph's
   *  style (minus structural markers); required when the reference block is not
   *  a paragraph. */
  style?: ParaStyle;
  /** Character style for the inserted text run. Defaults to the reference
   *  paragraph's first run style; required when neither is available. */
  runStyle?: CharStyle;
}

export class DocumentEditor {
  private _doc: Document;
  private readonly undoStack: Commit[] = [];
  private readonly redoStack: Commit[] = [];

  constructor(doc: Document) {
    this._doc = doc;
  }

  /** The current document. Every edit swaps this for a new immutable value
   *  (structural sharing), so a captured reference is a stable snapshot. */
  get doc(): Document {
    return this._doc;
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  // --- editing ----------------------------------------------------------------

  /** Replace all runs of a paragraph with a single run of `text`. `style`
   *  defaults to the paragraph's current first-run style. */
  setParagraphText(blockId: string, text: string, style?: CharStyle): this {
    const para = this.mustParagraph(blockId);
    const runs: Run[] = [];
    if (text.length > 0) {
      const runStyle = style ?? para.runs[0]?.style;
      if (!runStyle) throw new Error(`setParagraphText: paragraph ${blockId} is empty; pass a style`);
      runs.push({ text, style: runStyle });
    }
    return this.commit([{ type: "setRuns", blockId, runs }]);
  }

  /** Insert `text` at a UTF-16 offset within a paragraph. */
  insertText(blockId: string, offset: number, text: string, style?: CharStyle): this {
    return this.commit([{ type: "insertText", at: { blockId, offset }, text, ...(style ? { style } : {}) }]);
  }

  /** Delete the UTF-16 range [start, end) within a paragraph. */
  deleteText(blockId: string, start: number, end: number): this {
    return this.commit([{ type: "deleteRange", blockId, start, end }]);
  }

  /** Replace the UTF-16 range [start, end) within a paragraph with `text`
   *  (delete + insert as one undoable step). */
  replaceText(blockId: string, start: number, end: number, text: string, style?: CharStyle): this {
    const ops: Op[] = [{ type: "deleteRange", blockId, start, end }];
    if (text.length > 0) ops.push({ type: "insertText", at: { blockId, offset: start }, text, ...(style ? { style } : {}) });
    return this.commit(ops);
  }

  /** Patch a paragraph's style (alignment, indents, spacing, list, borders…). */
  setParagraphStyle(blockId: string, patch: Partial<ParaStyle>): this {
    return this.commit([{ type: "setParaStyle", blockId, patch }]);
  }

  /** Remove a top-level block (paragraph, table, image…) by id. */
  removeBlock(blockId: string): this {
    return this.commit([{ type: "removeBlock", blockId }]);
  }

  /** Insert a new paragraph before/after a top-level reference block. Returns the
   *  new paragraph's id via `lastInsertedId`. */
  insertParagraph(refBlockId: string, text: string, options: InsertParagraphOptions = {}): this {
    const found = containerOf(this._doc, refBlockId);
    if (!found) throw new Error(`insertParagraph: ${refBlockId} is not a top-level block`);
    const block = getBlockById(this._doc, refBlockId);
    const refPara = block?.kind === "paragraph" ? block : undefined;
    const style = options.style ?? (refPara ? cloneInsertableStyle(refPara.style) : undefined);
    if (!style) throw new Error(`insertParagraph: reference ${refBlockId} is not a paragraph; pass options.style`);
    const runStyle = options.runStyle ?? refPara?.runs[0]?.style;
    if (text.length > 0 && !runStyle) throw new Error(`insertParagraph: no run style available; pass options.runStyle`);
    const id = freshId();
    const paragraph: Paragraph = {
      kind: "paragraph",
      id,
      revision: 0,
      runs: text.length > 0 ? [{ text, style: runStyle! }] : [],
      style,
    };
    const index = options.position === "before" ? found.index : found.index + 1;
    this._lastInsertedId = id;
    return this.commit([{ type: "insertBlock", index, block: paragraph, where: found.where }]);
  }

  private _lastInsertedId: string | null = null;

  /** The id minted by the most recent `insertParagraph`, or null. */
  get lastInsertedId(): string | null {
    return this._lastInsertedId;
  }

  // --- content controls (SDTs) — the primary templating surface ---------------

  /** Merge a patch onto a content control's properties (alias, tag, checked,
   *  listItems, date format, locks…). The control must already exist. The control
   *  `type` is the discriminant for how it round-trips and paints, so it is NOT
   *  patchable here — it is always preserved (excluded from the patch shape and
   *  re-forced on the merge, even against an untyped runtime caller). */
  setSdtProps(id: string, patch: Partial<Omit<SdtProps, "type">>): this {
    const current = getSdt(this._doc, id);
    if (!current) throw new Error(`setSdtProps: content control ${id} not found`);
    const props: SdtProps = { ...current, ...patch, type: current.type };
    return this.commit([{ type: "setSdtProps", id, props }]);
  }

  /** Set a checkbox content control's state (☒/☐). Throws if `id` is not a
   *  checkbox control. */
  setCheckbox(id: string, checked: boolean): this {
    const current = getSdt(this._doc, id);
    if (!current) throw new Error(`setCheckbox: content control ${id} not found`);
    if (current.type !== "checkbox") {
      throw new Error(`setCheckbox: ${id} is a "${current.type}" control, not a checkbox`);
    }
    return this.setSdtProps(id, { checked });
  }

  /** Replace the text of a content control that occupies a SINGLE paragraph —
   *  either an inline control (runs tagged with `id`) or a block-level control
   *  wrapping one paragraph. The killer templating primitive: "fill this field".
   *  Preserves the control's ancestry (so nesting survives) and clears any
   *  placeholder flag. Throws for controls spanning multiple blocks (edit those by
   *  block id) or not found. */
  setSdtText(id: string, text: string): this {
    const current = getSdt(this._doc, id);
    if (!current) throw new Error(`setSdtText: content control ${id} not found`);

    const blocks = getSdtBlocks(this._doc, id);
    if (blocks.length > 0) {
      // Block-level control: require exactly one paragraph member.
      if (blocks.length > 1 || blocks[0]!.kind !== "paragraph") {
        throw new Error(`setSdtText: ${id} spans ${blocks.length} block(s); edit multi-block controls by block id`);
      }
      const para = blocks[0] as Paragraph;
      const runStyle = para.runs[0]?.style;
      if (text.length > 0 && !runStyle) {
        throw new Error(`setSdtText: block control ${id} is empty; cannot infer a run style`);
      }
      // setRuns keeps Block.sdtPath (block-level membership) intact.
      const runs: Run[] = text.length > 0 ? [{ text, style: runStyle! }] : [];
      return this.commit(this.withPlaceholderCleared(id, current, [{ type: "setRuns", blockId: para.id, runs }]));
    }

    // Inline control: find the single paragraph carrying tagged runs.
    const hit = this.locateInlineSdt(id);
    if (!hit) throw new Error(`setSdtText: content control ${id} has no editable text region`);
    const { para, from, to, baseStyle } = hit;
    // Preserve the run's full ancestry up to and including `id` so outer controls
    // survive and the new content sits at this control's level.
    const style: CharStyle = { ...baseStyle, sdtPath: ancestryThrough(baseStyle.sdtPath, id) ?? [id] };
    const rebuilt: Run[] = [
      ...para.runs.slice(0, from),
      ...(text.length > 0 ? [{ text, style }] : []),
      ...para.runs.slice(to),
    ];
    return this.commit(this.withPlaceholderCleared(id, current, [{ type: "setRuns", blockId: para.id, runs: rebuilt }]));
  }

  /** Select a value on a dropDown/comboBox content control — the "choose an
   *  option" primitive. `value` matches a listItem by its VALUE first, then by its
   *  display; the control's text becomes that item's display. A comboBox accepts
   *  free text (no match ⇒ `value` is used verbatim); a dropDown requires a listed
   *  option (throws otherwise). For other control kinds use `setSdtText`/`setCheckbox`. */
  setSdtValue(id: string, value: string): this {
    const props = getSdt(this._doc, id);
    if (!props) throw new Error(`setSdtValue: content control ${id} not found`);
    if (props.type !== "dropDown" && props.type !== "comboBox") {
      throw new Error(`setSdtValue: ${id} is a "${props.type}" control; use setSdtText or setCheckbox`);
    }
    const items = props.listItems ?? [];
    const match = items.find((i) => i.value === value) ?? items.find((i) => i.display === value);
    if (!match && props.type === "dropDown") {
      throw new Error(`setSdtValue: "${value}" is not an option of dropdown ${id}`);
    }
    return this.setSdtText(id, match ? match.display : value);
  }

  /** Remove a content control, UNWRAPPING it: strip its id from every member run
   *  (inline) and body block (block-level) path — so NESTED controls survive — and
   *  delete its properties, keeping the "every path id has a props entry" invariant.
   *  With `deleteContents: true` the wrapped content is deleted too. Inline controls
   *  are handled in any story (body, cells, bands, notes); block-level controls are
   *  handled at the TOP LEVEL of the body — a block-level control nested in a table
   *  cell, band, or note throws (remove it by block id). */
  removeSdt(id: string, options: { deleteContents?: boolean } = {}): this {
    const props = getSdt(this._doc, id);
    if (!props) throw new Error(`removeSdt: content control ${id} not found`);
    const deleteContents = options.deleteContents ?? false;
    const bodyIndex = new Map(this._doc.blocks.map((b, i) => [b.id, i] as const));

    // Gather members: inline paragraphs (a run carries id) and block-level members.
    const inlineParaIds: string[] = [];
    const blockMembers: Block[] = [];
    let unsupported: string | undefined;
    walk(this._doc, (b, ctx) => {
      if (b.sdtPath?.includes(id)) {
        if (ctx.container === "body" && !ctx.cell && !ctx.note && bodyIndex.has(b.id)) blockMembers.push(b);
        else unsupported ??= ctx.note ? "a note body" : ctx.cell ? "a table cell" : `the ${ctx.container} band`;
      } else if (b.kind === "paragraph" && b.runs.some((r) => r.style.sdtPath?.includes(id))) {
        inlineParaIds.push(b.id);
      }
    });
    if (unsupported) {
      throw new Error(`removeSdt: ${id} has a block-level member in ${unsupported}; not supported — remove it by block id`);
    }

    const ops: Op[] = [];
    // Inline members: rewrite each paragraph's runs (strip id, or drop tagged runs).
    for (const blockId of inlineParaIds) {
      const para = getParagraphById(this._doc, blockId)!;
      const runs = deleteContents
        ? para.runs.filter((r) => !r.style.sdtPath?.includes(id))
        : para.runs.map((r) => (r.style.sdtPath?.includes(id) ? stripRunSdt(r, id) : r));
      ops.push({ type: "setRuns", blockId, runs });
    }
    // Block-level members (top-level body): remove, then re-insert stripped unless
    // deleting content. Descending index so each original index stays valid.
    for (const b of blockMembers.slice().sort((a, z) => bodyIndex.get(z.id)! - bodyIndex.get(a.id)!)) {
      const index = bodyIndex.get(b.id)!;
      ops.push({ type: "removeBlock", blockId: b.id });
      if (!deleteContents) ops.push({ type: "insertBlock", index, block: stripBlockSdt(b, id), where: "body" });
    }
    ops.push({ type: "setSdtProps", id, props: null });
    return this.commit(ops);
  }

  // --- ergonomic bulk / structural edits --------------------------------------

  /** Find/replace across every paragraph in the document as ONE undoable step.
   *  A string replaces ALL occurrences; a RegExp is applied per run (pass the `g`
   *  flag to replace all). Replacement happens WITHIN each run — matches that span
   *  a style boundary are not replaced (a deliberate limit so run styling is always
   *  preserved). Returns `this`. */
  replaceAllText(pattern: string | RegExp, replacement: string): this {
    const ops: Op[] = [];
    walk(this._doc, (b) => {
      if (b.kind !== "paragraph") return;
      let changed = false;
      const runs = b.runs.map((r) => {
        const text = typeof pattern === "string" ? r.text.split(pattern).join(replacement) : r.text.replace(pattern, replacement);
        if (text !== r.text) changed = true;
        return text === r.text ? r : { ...r, text };
      });
      if (changed) ops.push({ type: "setRuns", blockId: b.id, runs });
    });
    return this.commit(ops);
  }

  /** Apply a named paragraph style by its human NAME (e.g. "Heading 1"), resolving
   *  it to a styleId via the stylesheet and BAKING the resolved paragraph + run
   *  formatting onto the block (the model is concrete), plus setting the `namedStyle`
   *  reference. Throws if the document has no stylesheet or no style with that name. */
  setStyleByName(blockId: string, styleName: string): this {
    const sheet = this._doc.stylesheet;
    const style = sheet?.styles.find((s) => s.name === styleName);
    if (!sheet || !style) throw new Error(`setStyleByName: no style named "${styleName}"`);
    const para = this.mustParagraph(blockId);
    const resolved = resolveStyle(sheet, style.id);
    const ops: Op[] = [{ type: "setParaStyle", blockId, patch: { ...resolved.para, namedStyle: style.id } }];
    if (Object.keys(resolved.char).length > 0 && para.runs.length > 0) {
      const len = para.runs.reduce((n, r) => n + r.text.length, 0);
      ops.push({ type: "setRuns", blockId, runs: applyStylePatchToRuns(para.runs, 0, len, resolved.char) });
    }
    return this.commit(ops);
  }

  /** Move a top-level BODY block to a new index among the body blocks (remove +
   *  insert as one undoable step). `toIndex` is the destination among the OTHER
   *  blocks (after removal), clamped to range. Throws if the block is not a
   *  top-level body block. */
  moveBlock(blockId: string, toIndex: number): this {
    const from = this._doc.blocks.findIndex((b) => b.id === blockId);
    if (from === -1) throw new Error(`moveBlock: ${blockId} is not a top-level body block`);
    const block = this._doc.blocks[from]!;
    const dest = Math.max(0, Math.min(toIndex, this._doc.blocks.length - 1));
    if (dest === from) return this;
    return this.commit([
      { type: "removeBlock", blockId },
      { type: "insertBlock", index: dest, block, where: "body" },
    ]);
  }

  /** Insert a row into a table at `rowIndex` (clamped). `cellTexts` fills the new
   *  cells left-to-right (missing/extra entries pad/trim to the table's column
   *  count); omit for an empty row. New cells inherit a run style from an existing
   *  table paragraph (or the document default). */
  insertTableRowAt(tableId: string, rowIndex: number, cellTexts: string[] = []): this {
    const table = getTableById(this._doc, tableId);
    if (!table) throw new Error(`insertTableRowAt: table ${tableId} not found`);
    // Span-aware column count so a first row with colSpans doesn't undercount;
    // `gridColumnCount` never returns 0, so for a brand-new EMPTY table fall back
    // to the requested cell count (else it would force a single column).
    const cols = table.rows.length > 0 ? gridColumnCount(table) : Math.max(1, cellTexts.length);
    const runStyle = firstRunStyle(table.rows) ?? DEFAULT_CHAR_STYLE;
    const cells: TableCell[] = Array.from({ length: cols }, (_, c) => {
      const text = cellTexts[c] ?? "";
      const paragraph: Paragraph = {
        kind: "paragraph",
        id: freshId(),
        revision: 0,
        runs: text.length > 0 ? [{ text, style: runStyle }] : [],
        style: { ...DEFAULT_PARA_STYLE },
      };
      return { id: freshId(), blocks: [paragraph] };
    });
    const at = Math.max(0, Math.min(rowIndex, table.rows.length));
    return this.commit([{ type: "insertTableRow", tableId, rowIndex: at, row: { cells } }]);
  }

  /** Delete the column whose header (first-row cell text) equals `headerText`.
   *  Throws if the table or a matching header is not found. */
  deleteColumnByHeader(tableId: string, headerText: string): this {
    const table = getTableById(this._doc, tableId);
    if (!table) throw new Error(`deleteColumnByHeader: table ${tableId} not found`);
    const header = table.rows[0];
    const col = header ? header.cells.findIndex((cell) => textOfCell(cell) === headerText) : -1;
    if (col === -1) throw new Error(`deleteColumnByHeader: no column headed "${headerText}" in table ${tableId}`);
    return this.commit([{ type: "removeTableColumn", tableId, colIndex: col }]);
  }

  /** Find the single paragraph holding the contiguous run span tagged with the
   *  inline control `id`, plus that span `[from, to)` and its base run style. */
  private locateInlineSdt(
    id: string,
  ): { para: Paragraph; from: number; to: number; baseStyle: CharStyle } | undefined {
    let found: { para: Paragraph; from: number; to: number; baseStyle: CharStyle } | undefined;
    walk(this._doc, (block) => {
      if (found || block.kind !== "paragraph") return;
      let from = -1;
      let to = -1;
      block.runs.forEach((run, i) => {
        if (run.style.sdtPath?.includes(id)) {
          if (from === -1) from = i;
          to = i + 1;
        }
      });
      if (from !== -1) found = { para: block, from, to, baseStyle: block.runs[from]!.style };
    });
    return found;
  }

  /** Append a `setSdtProps` op clearing a placeholder flag, when set — so filling
   *  a control drops its gray placeholder styling, like the interactive editor. */
  private withPlaceholderCleared(id: string, current: SdtProps, ops: Op[]): Op[] {
    if (!current.placeholder) return ops;
    return [...ops, { type: "setSdtProps", id, props: { ...current, placeholder: false } }];
  }

  // --- undo / redo ------------------------------------------------------------

  /** Undo the most recent edit. Returns false when there's nothing to undo. */
  undo(): boolean {
    const commit = this.undoStack.pop();
    if (!commit) return false;
    let doc = this._doc;
    for (const op of commit.inverse) doc = applyOp(doc, op).doc;
    this._doc = doc;
    this.redoStack.push(commit);
    return true;
  }

  /** Redo the most recently undone edit. Returns false when there's nothing to redo. */
  redo(): boolean {
    const commit = this.redoStack.pop();
    if (!commit) return false;
    let doc = this._doc;
    for (const op of commit.forward) doc = applyOp(doc, op).doc;
    this._doc = doc;
    this.undoStack.push(commit);
    return true;
  }

  /** Append another document's content after this one — the headless equivalent of
   *  Word's "insert file at end". Reconciles every id space (see mergeDocuments) and
   *  lands as ONE undoable step. Returns the merge's id map + warnings; the merged
   *  document becomes `this.doc`. */
  append(source: Document, options: MergeOptions = {}): MergeResult {
    const result = mergeDocuments(this._doc, source, options);
    this.commit([{ type: "setDocument", doc: result.doc }]);
    return result;
  }

  // --- convenience queries ----------------------------------------------------

  /** Find paragraphs by text (see query.findParagraphs). */
  find(pattern: string | RegExp): ParagraphMatch[] {
    return findParagraphs(this._doc, pattern);
  }

  /** The paragraph with this id, or undefined. */
  getParagraph(blockId: string): Paragraph | undefined {
    return getParagraphById(this._doc, blockId);
  }

  // --- internals --------------------------------------------------------------

  private mustParagraph(blockId: string): Paragraph {
    const para = getParagraphById(this._doc, blockId);
    if (!para) throw new Error(`paragraph ${blockId} not found`);
    return para;
  }

  /** Apply a list of ops as ONE undoable commit, clearing the redo stack. */
  private commit(ops: Op[]): this {
    if (ops.length === 0) return this;
    const inverse: Op[] = [];
    let doc = this._doc;
    for (const op of ops) {
      const result = applyOp(doc, op);
      doc = result.doc;
      inverse.unshift(result.inverse); // reverse-apply order for undo
    }
    this._doc = doc;
    this.undoStack.push({ forward: ops, inverse });
    this.redoStack.length = 0;
    return this;
  }
}

function cloneInsertableStyle(style: ParaStyle): ParaStyle {
  const clone: ParaStyle = { ...style };
  const bag = clone as unknown as Record<string, unknown>;
  for (const key of NON_INHERITED_STYLE_KEYS) delete bag[key];
  return clone;
}

/** A run with control `id` removed from its inline sdt path (nesting preserved);
 *  the marker is dropped entirely once the path becomes empty. */
function stripRunSdt(run: Run, id: string): Run {
  const path = stripSdtFromPath(run.style.sdtPath, id);
  const style: CharStyle = { ...run.style };
  if (path) style.sdtPath = path;
  else delete style.sdtPath;
  return { ...run, style };
}

/** A block with control `id` removed from its block-level sdt path (nesting
 *  preserved); revision bumped so the layout/prepare caches invalidate. */
function stripBlockSdt(block: Block, id: string): Block {
  const path = stripSdtFromPath(block.sdtPath, id);
  const next = { ...block, revision: block.revision + 1 } as Block;
  const bag = next as { sdtPath?: string[] };
  if (path) bag.sdtPath = path;
  else delete bag.sdtPath;
  return next;
}

/** The style of the first run found in any cell of `rows` — a template for new
 *  cell paragraphs. Undefined if every cell is empty. */
function firstRunStyle(rows: TableRow[]): CharStyle | undefined {
  for (const row of rows) {
    for (const cell of row.cells) {
      for (const block of cell.blocks) {
        if (block.kind === "paragraph" && block.runs[0]) return block.runs[0].style;
      }
    }
  }
  return undefined;
}
