// Undo manager: stacks of op groups with their inverses. Consecutive "typing"
// transactions coalesce into one undo step, broken by a >1s pause or any
// non-typing transaction (Word behavior). Transient (IME preview) transactions
// never reach this module.

import type { Op } from "@cw/shared";
import type { DocSelection } from "@cw/shared";
import type { TransactionOrigin } from "./state";

export interface UndoEntry {
  /** Forward ops (for redo), in application order. */
  ops: Op[];
  /** Inverse ops (for undo), already in reverse application order. */
  inverseOps: Op[];
  selectionBefore: DocSelection | null;
  selectionAfter: DocSelection | null;
  origin: TransactionOrigin;
  time: number;
}

const COALESCE_MS = 1000;

export class UndoManager {
  private undoStack: UndoEntry[] = [];
  private redoStack: UndoEntry[] = [];

  record(entry: UndoEntry): void {
    this.redoStack.length = 0;
    const last = this.undoStack[this.undoStack.length - 1];
    if (
      last &&
      last.origin === "typing" &&
      entry.origin === "typing" &&
      entry.time - last.time < COALESCE_MS
    ) {
      last.ops.push(...entry.ops);
      last.inverseOps = [...entry.inverseOps, ...last.inverseOps];
      last.selectionAfter = entry.selectionAfter;
      last.time = entry.time;
      return;
    }
    this.undoStack.push(entry);
  }

  popUndo(): UndoEntry | null {
    const entry = this.undoStack.pop();
    if (!entry) return null;
    this.redoStack.push(entry);
    return entry;
  }

  popRedo(): UndoEntry | null {
    const entry = this.redoStack.pop();
    if (!entry) return null;
    this.undoStack.push(entry);
    return entry;
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }
  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }
}
