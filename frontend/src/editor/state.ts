// Layer 5: editor core. One-way data flow:
// input -> command -> transaction -> applyOp* -> new EditorState -> layout -> paint (rAF).

import type { CharStyle, Document } from "@cw/shared";
import type { DocSelection } from "@cw/shared";
import type { Op } from "@cw/shared";

/** A rectangular block of table cells selected by dragging across them, in GRID
 *  coordinates (span-aware; see shared/tableGrid). Drives merge and the Borders &
 *  Shading editor. Distinct from the text `selection` — only one is active at a
 *  time. Body tables only for now. */
export interface CellSelection {
  tableId: string;
  anchor: { row: number; col: number };
  focus: { row: number; col: number };
}

export interface EditorState {
  doc: Document;
  selection: DocSelection | null;
  /** Active rectangular table-cell selection, if any (suppresses text selection). */
  cellSelection?: CellSelection | null;
  /** Style toggled at a collapsed caret — applied to the NEXT typed text and
   *  dropped when the caret moves (Word behavior). */
  pendingStyle?: Partial<CharStyle> | null;
}

export type TransactionOrigin =
  | "typing" // coalesces into one undo step
  | "command"
  | "paste"
  | "transient" // IME composition preview — bypasses the undo stack
  | "undo"
  | "redo";

export interface Transaction {
  ops: Op[];
  selectionAfter: DocSelection | null;
  origin: TransactionOrigin;
}

/** Commands are pure (state) -> Transaction | null. Keymap entries and toolbar
 *  buttons share this single code path, so editing is testable headless. */
export type Command = (state: EditorState) => Transaction | null;
