// Undo manager: stacks of op groups with their inverses. Consecutive "typing"
// transactions coalesce into one undo step, broken by a >1s pause or any
// non-typing transaction (Word behavior). Transient (IME preview) transactions
// never reach this module.

import type { Op } from "@cw/shared";
import type { DocSelection } from "@cw/shared";
import type { ReviewOp } from "@cw/shared";
import type { TransactionOrigin } from "./state";

export interface UndoEntry {
  /** Forward ops (for redo), in application order. */
  ops: Op[];
  /** Inverse ops (for undo), already in reverse application order. */
  inverseOps: Op[];
  /** Review-layer forward ops this action also applied (suggest mode), in
   *  application order — replayed on redo. Empty for plain text edits. */
  reviewOps?: ReviewOp[];
  /** Review-layer inverses, in reverse application order — replayed on undo,
   *  AFTER the text inverses (so anchors are valid when records are restored). */
  reviewInverses?: ReviewOp[];
  selectionBefore: DocSelection | null;
  selectionAfter: DocSelection | null;
  origin: TransactionOrigin;
  time: number;
}

const COALESCE_MS = 1000;

export class UndoManager {
  private undoStack: UndoEntry[] = [];
  private redoStack: UndoEntry[] = [];
  // While a typing run coalesces, its accumulated inverses live here in
  // APPLICATION order, so each keystroke is an O(its own ops) append — the old
  // prepend-and-respread (`[...entry.inverseOps, ...last.inverseOps]`) copied
  // the whole run per keystroke, O(n²) array churn across a long burst. They
  // fold back into the entry, in the documented reverse-application order, the
  // moment the run closes (a non-coalescing record, or a pop).
  private runInverses: Op[] | null = null;
  private runReviewInverses: ReviewOp[] | null = null;

  /** Fold an open typing run's accumulated inverses back onto its entry. */
  private closeRun(): void {
    if (!this.runInverses && !this.runReviewInverses) return;
    const last = this.undoStack[this.undoStack.length - 1];
    if (last && this.runInverses) last.inverseOps = this.runInverses.reverse();
    if (last && this.runReviewInverses) last.reviewInverses = this.runReviewInverses.reverse();
    this.runInverses = null;
    this.runReviewInverses = null;
  }

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
      // entry.inverseOps is reverse-ordered within the entry; the run buffer is
      // application-ordered, so append the entry's inverses reversed.
      this.runInverses ??= last.inverseOps.slice().reverse();
      for (let i = entry.inverseOps.length - 1; i >= 0; i--) this.runInverses.push(entry.inverseOps[i]!);
      if (entry.reviewOps?.length) {
        if (!last.reviewOps) last.reviewOps = [];
        last.reviewOps.push(...entry.reviewOps);
      }
      if (entry.reviewInverses?.length) {
        this.runReviewInverses ??= (last.reviewInverses ?? []).slice().reverse();
        for (let i = entry.reviewInverses.length - 1; i >= 0; i--) this.runReviewInverses.push(entry.reviewInverses[i]!);
      }
      last.selectionAfter = entry.selectionAfter;
      last.time = entry.time;
      return;
    }
    this.closeRun();
    this.undoStack.push(entry);
  }

  popUndo(): UndoEntry | null {
    this.closeRun(); // the popped entry may be the still-open typing run
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
