// DocumentEditor — an ergonomic, headless editing facade over the `applyOp`
// operation engine, the rough analog of mutating a .NET `WordprocessingDocument`
// and saving. It holds a mutable `doc`, translates high-level calls into the
// existing typed `Op`s, threads them through `applyOp` (which does structural
// sharing and produces exact inverses), and keeps an undo/redo stack for free.
//
// This is deliberately NOT the interactive editor's Command/Transaction/caret
// machinery (that lives in the frontend and is UI-coupled). It is a plain data
// facade usable from Node, the browser, and — later — the C# bindings.

import type { CharStyle, Document, Paragraph, ParaStyle, Run } from "./document";
import { applyOp, containerOf, type Op } from "./ops";
import { findParagraphs, getBlockById, getParagraphById, type ParagraphMatch } from "./query";
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
    const runStyle = style ?? para.runs[0]?.style;
    if (!runStyle) throw new Error(`setParagraphText: paragraph ${blockId} is empty; pass a style`);
    const runs: Run[] = text.length > 0 ? [{ text, style: runStyle }] : [];
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
