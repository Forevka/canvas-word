// Mouse/keyboard selection over the canvas, driven entirely by layout hit-testing.
//
// - mousedown -> hitTest -> anchor; drag extends focus; auto-scroll near edges.
// - double-click: word select via Intl.Segmenter (word granularity) — the same
//   segmentation pretext uses, so selections align with break opportunities.
// - triple-click: paragraph select.
// - Arrows move by grapheme cluster, Ctrl+arrow by word, Home/End by line
//   (a layout query, not a model one), Up/Down preserve goalX, shift extends,
//   Ctrl+A selects all, Ctrl+C copies plain text (HTML flavor lands in milestone 5).

import type { Document, ImageBlock, Paragraph } from "@cw/shared";
import type { DocPosition, DocSelection } from "@cw/shared";
import { isCollapsed } from "@cw/shared";
import {
  words,
  blockById,
  navigableBandParagraphs,
  navigableParagraphs,
  prevGrapheme,
  nextGrapheme,
  prevWordStart,
  nextWordEnd,
  type BandName,
} from "@cw/shared";
import { findTableById } from "../editor/commands";
import type { LayoutTree } from "../layout/layoutTree";
import {
  comparePositions,
  documentEdges,
  hitTest,
  hitTestCell,
  hitTestSelectableObject,
  hitTestColumnBoundary,
  hitTestRowBoundary,
  inlineEquationAt,
  caretRect,
  lineEdges,
  linkAt,
  objectRect,
  positionOnAdjacentLine,
  visualCaretStep,
  type ColumnBoundaryHit,
  type RowBoundaryHit,
  type GeoScope,
  type BandScope,
} from "../layout/geometry";
import type { ColumnGuide, RowGuide, PagePoint } from "../paint/renderer";
import type { CellSelection } from "../editor/state";
import { extractFragment, fragmentToHtml, fragmentToPlainText, tableRectToClipboard } from "./clipboard";

/** Unit nudge direction (x, y in document px, y-down) per arrow key — the E2
 *  object-nudge multiplies these by 1px (or 10px with Shift). */
const ARROW_DELTA: Record<string, [number, number] | undefined> = {
  ArrowLeft: [-1, 0],
  ArrowRight: [1, 0],
  ArrowUp: [0, -1],
  ArrowDown: [0, 1],
};

export interface SelectionControllerDeps {
  container: HTMLElement;
  getTree(): LayoutTree;
  getDoc(): Document;
  getSelection(): DocSelection | null;
  setSelection(sel: DocSelection | null): void;
  /** The active rectangular table-cell selection, if any (for copy/cut). */
  getCellSelection(): CellSelection | null;
  /** Set (or clear) the rectangular table-cell selection during a cross-cell drag. */
  setCellSelection(sel: CellSelection | null): void;
  clientToPage(clientX: number, clientY: number): PagePoint | null;
  /** Drawing-grid step in document px; 0 disables snapping. */
  getGridSpacing(): number;
  /** Whether an anchored-object drag should snap to the grid. */
  isSnapToGrid(): boolean;
  /** Route focus to the IME proxy so typing works right after a click. */
  focusProxy(): void;
  /** Delete the current selection (Ctrl+X after the copy half). */
  onDeleteSelection(): void;
  /** Story-edit scope (header/footer mode). Null = editing the body. */
  getStory(): BandScope | null;
  /** Enter (band+page) or exit (null) story mode; triggers relayout + dimming. */
  setStory(scope: BandScope | null): void;
  /** An interactive embedder decoration is at this client point (its source-spec
   *  index), or null. Drives the hover cursor. */
  decorationAt(clientX: number, clientY: number): number | null;
  /** Fire the click handler of an interactive decoration at this client point.
   *  Returns true when one handled it (the click is then consumed). */
  onDecorationClick(clientX: number, clientY: number): boolean;
  /** Select an image object (null clears). */
  selectObject(blockId: string | null): void;
  /** True when an object is selected (Delete/Backspace/Escape route to it). */
  hasSelectedObject(): boolean;
  /** Move the object selection to the next (or previous, `backward`) selectable
   *  object in document order, wrapping; from nothing selected it enters at the
   *  first/last. The keyboard object-focus cycle (Tab while an object is selected,
   *  and F6 from anywhere — E3). Returns false only when there are no objects. */
  cycleObjectFocus(backward: boolean): boolean;
  /** Arrow-key nudge (E2): move the selected anchored shape by `dxPx`/`dyPx`
   *  (page-clamped). Returns true when the selection is a nudgeable anchored shape
   *  (the arrow is then consumed), false so the arrow falls through otherwise. */
  nudgeSelectedObject(dxPx: number, dyPx: number): boolean;
  deleteSelectedObject(): void;
  /** Begin a table column-boundary drag (wiring owns the transient/commit loop). */
  startColumnDrag(hit: ColumnBoundaryHit, ev: MouseEvent): void;
  /** Highlight (or clear) the column boundary under the pointer before a drag. */
  setColumnGuide(guide: ColumnGuide | null): void;
  /** Begin a table row-boundary drag (wiring owns the transient/commit loop). */
  startRowDrag(hit: RowBoundaryHit, ev: MouseEvent): void;
  /** Highlight (or clear) the row boundary under the pointer before a drag. */
  setRowGuide(guide: RowGuide | null): void;
  /** Reposition an anchored image during a drag. `transient` = live preview
   *  (outside undo); false = the committed drop (one undoable step). */
  applyObjectMove(blockId: string, offsetXPx: number, offsetYPx: number, transient: boolean): void;
  /** Tab/Shift+Tab cell navigation; returns true when consumed. */
  onTab(backward: boolean): boolean;
  /** Move the caret to a block's start and scroll it into view (TOC jump). */
  jumpToBlock(blockId: string): void;
  /** Ctrl+click on an in-document anchor link (#bookmark). `fromBlockId` is the
   *  clicked paragraph (used to resolve the target by text when the bookmark
   *  isn't modeled). The wiring scrolls to the resolved heading. */
  onAnchorJump(anchorName: string, fromBlockId: string | null): void;
  /** Single click landed on a content control. Return true to CONSUME the
   *  press (checkbox toggle); false lets the caret place normally (dropdown /
   *  date popups open beside the caret). */
  onSdtPress(pos: DocPosition): boolean;

  // ---- shape text box editing (issue #219) ------------------------------
  /** The drawing shape whose text box body is being caret-edited, or null. */
  getActiveShapeText(): string | null;
  /** The geometry scope for the active shape text box (null when not editing). */
  getShapeScope(): GeoScope | null;
  /** The shape's text box paragraphs, in order (the box's closed nav sub-flow). */
  getShapeTextParagraphs(shapeId: string): Paragraph[];
  /** The shape's placed box rect (for in-box hit routing), or null. */
  shapeTextRect(shapeId: string): { pageIndex: number; x: number; y: number; width: number; height: number } | null;
  /** Does this shape carry an editable text box body? */
  shapeHasText(shapeId: string): boolean;
  /** Enter a shape's text box, placing the caret at `pos` inside it. */
  enterShapeText(shapeId: string, pos: DocPosition): void;
  /** Leave the active text box: `toObject` pops back to selecting the shape
   *  (Escape); false just clears the session (a click outside commits). */
  exitShapeText(toObject: boolean): void;
  /** Enter the currently-selected shape's text box (Enter on a selected shape).
   *  Returns true when it did. */
  enterSelectedShapeText(): boolean;
  /** Start a text body on a TOP-LEVEL, text-less shape and enter it (double-click
   *  "Add text"). Returns false for a cell-nested shape (stays read-only) or a
   *  non-editable context, so the caller falls back to plain object selection. #235. */
  startShapeText(shapeId: string): boolean;
}

export interface SelectionController {
  destroy(): void;
}

/** A painted caret point — the page it lands on and its position there. */
export interface CaretXY {
  page: number;
  x: number;
  y: number;
}

/** Step `offset` within ONE paragraph's `text` by a grapheme (or word) in `dir`,
 *  then keep stepping while the new offset paints at the SAME caret point as the
 *  start. A {page}/{pages} token field collapses its whole model range onto one
 *  glyph (offsetMap) and a hidden (w:vanish) run lays out zero-width, so several
 *  model offsets map onto ONE visual point — a plain step into that span reads as
 *  a dead arrow key (the caret only moves once you've pressed past the run). This
 *  collapses those offsets into a single stop, so one press always moves the caret
 *  (Word does the same). Returns the resting offset, or null when the walk reaches
 *  the block edge while still collapsed (the caller then crosses to the sibling). */
export function advanceVisibleStop(
  text: string,
  offset: number,
  dir: -1 | 1,
  byWord: boolean,
  caretXY: (offset: number) => CaretXY | null,
): number | null {
  const from = caretXY(offset);
  const collapsed = (off: number): boolean => {
    const r = caretXY(off);
    return !!from && !!r && from.page === r.page && Math.abs(from.x - r.x) < 0.5 && Math.abs(from.y - r.y) < 0.5;
  };
  let o = offset;
  if (dir === -1) {
    while (o > 0) {
      o = byWord ? prevWordStart(text, o) : prevGrapheme(text, o);
      if (!collapsed(o)) return o;
    }
    return null;
  }
  while (o < text.length) {
    o = byWord ? nextWordEnd(text, o) : nextGrapheme(text, o);
    if (!collapsed(o)) return o;
  }
  return null;
}

export function createSelectionController(deps: SelectionControllerDeps): SelectionController {
  const { container } = deps;
  container.tabIndex = 0;
  container.style.outline = "none";
  container.style.cursor = "text";

  // ---- model text helpers ----------------------------------------------

  const scope = (): GeoScope | undefined => deps.getStory() ?? deps.getShapeScope() ?? undefined;

  // Navigation-order paragraphs for the ACTIVE story: while story-editing,
  // the variant CONTAINER that renders on the story's page (first/even bands
  // override the default); otherwise the body (including table cells).
  const paragraphs = (): Paragraph[] => {
    // Editing a shape text box: the box is a CLOSED sub-flow — the caret navigates
    // only its own paragraphs and stops at the box edges (#219).
    const shapeId = deps.getActiveShapeText();
    if (shapeId) return deps.getShapeTextParagraphs(shapeId);
    const story = deps.getStory();
    if (story) {
      const pg = deps.getTree().pages[story.pageIndex];
      const source =
        (story.band === "header" ? pg?.headerSource : pg?.footerSource) ?? story.band;
      return navigableBandParagraphs(deps.getDoc(), source);
    }
    // Hidden (w:vanish) paragraphs aren't laid out, so the caret must skip them.
    // Memoized per document identity (rebuilt only when the doc changes).
    return navigableParagraphs(deps.getDoc());
  };

  const textOf = (blockId: string): string => {
    const block = paragraphs().find((b) => b.id === blockId);
    return block ? block.runs.map((r) => r.text).join("") : "";
  };

  const blockIndexOf = (blockId: string): number =>
    paragraphs().findIndex((b) => b.id === blockId);

  // ---- caret movement primitives ----------------------------------------

  const moveHorizontal = (pos: DocPosition, dir: -1 | 1, byWord: boolean): DocPosition => {
    const text = textOf(pos.blockId);
    const blocks = paragraphs();
    const bi = blockIndexOf(pos.blockId);
    const caretXY = (offset: number): CaretXY | null => {
      const r = caretRect(deps.getTree(), { blockId: pos.blockId, offset }, scope());
      return r ? { page: r.pageIndex, x: r.x, y: r.y } : null;
    };
    const stop = advanceVisibleStop(text, pos.offset, dir, byWord, caretXY);
    if (stop !== null) return { blockId: pos.blockId, offset: stop };
    // The walk fell off the block edge (its whole prefix/suffix was collapsed):
    // cross into the adjacent paragraph, as a plain edge step would.
    if (dir === -1) {
      const prev = blocks[bi - 1];
      return prev ? { blockId: prev.id, offset: textOf(prev.id).length } : pos;
    }
    const next = blocks[bi + 1];
    return next ? { blockId: next.id, offset: 0 } : pos;
  };

  /** Visual-order horizontal step (bidi-aware). On a bidi line, ArrowLeft/Right
   *  move to the visually adjacent caret stop; at the line's visual edge (or on a
   *  plain LTR line, or for Ctrl+word steps) it falls back to the logical step,
   *  with the direction flipped inside an RTL-base paragraph so the caret crosses
   *  lines/blocks in the correct visual sense. */
  const moveHorizontalVisual = (pos: DocPosition, dir: -1 | 1, byWord: boolean): DocPosition => {
    if (!byWord) {
      const v = visualCaretStep(deps.getTree(), pos, dir, scope());
      if (v) return v;
    }
    const rtlBase = paragraphs().find((b) => b.id === pos.blockId)?.style.direction === "rtl";
    const logicalDir: -1 | 1 = rtlBase ? (dir === -1 ? 1 : -1) : dir;
    return moveHorizontal(pos, logicalDir, byWord);
  };

  const applyMove = (target: DocPosition, extend: boolean, goalX?: number): void => {
    const sel = deps.getSelection();
    const next: DocSelection = {
      anchor: extend && sel ? sel.anchor : target,
      focus: target,
    };
    if (goalX !== undefined) next.goalX = goalX;
    deps.setSelection(next);
  };

  // ---- mouse -------------------------------------------------------------

  let autoScrollDir = 0;
  let autoScrollRaf: number | null = null;
  // The active drag session, or null when idle. `cellTable`/`cellAnchor` are the
  // table + anchor cell captured on mousedown for a potential cross-cell drag
  // (crossing into another cell switches text selection → rectangular cell select).
  interface DragSession {
    cellTable: string | null;
    cellAnchor: { row: number; col: number } | null;
  }
  let drag: DragSession | null = null;

  // An in-progress drag of an anchored image. startX/startY are the page-space
  // grab point; baseX/baseY the image's anchor offsets at grab; lastX/lastY the
  // most recent previewed offsets (committed on mouseup). `moved` gates the first
  // few pixels so a plain click stays a select, not a (no-op, undoable) move.
  interface ObjectDrag {
    blockId: string;
    startX: number;
    startY: number;
    baseX: number;
    baseY: number;
    lastX: number;
    lastY: number;
    /** Absolute (page-local) origin of the anchor's reference frame, so snapping
     *  lands the placed object on the drawn grid regardless of relFrom*. */
    origin: { ox: number; oy: number };
    moved: boolean;
  }
  let objectDrag: ObjectDrag | null = null;

  /** The top-level anchored (out-of-flow, draggable) image OR shape with this id, if
   *  any — both carry the same anchor offset shape, so a drag repositions either. */
  const movableObject = (blockId: string): { anchor: NonNullable<ImageBlock["anchor"]> } | null => {
    const b = deps.getDoc().blocks.find((x) => x.id === blockId);
    return (b?.kind === "image" || b?.kind === "shape") && b.anchor ? { anchor: b.anchor } : null;
  };

  const posFromEvent = (ev: MouseEvent): DocPosition | null => {
    const pt = deps.clientToPage(ev.clientX, ev.clientY);
    if (!pt) return null;
    return hitTest(deps.getTree(), pt.pageIndex, pt.x, pt.y, scope());
  };

  /** Which margin band (if any) a page-local point falls in. */
  const bandAtPoint = (pt: PagePoint): BandName | null => {
    const tree = deps.getTree();
    const pg = tree.pages[pt.pageIndex];
    if (!pg) return null;
    // Presence is per PAGE: a first/even variant may exist where the default
    // band doesn't (and vice versa) — the tree records what rendered there.
    // Tall bands push the content box, so the clickable band regions are
    // everything OUTSIDE the body's real content box.
    if (pt.y < pg.contentTopPx) return pg.headerSource ? "header" : null;
    if (pt.y > pg.contentBottomPx) return pg.footerSource ? "footer" : null;
    return null;
  };

  const selectWordAt = (pos: DocPosition): void => {
    const text = textOf(pos.blockId);
    let start = 0;
    let end = text.length;
    for (const s of words.segment(text)) {
      const sEnd = s.index + s.segment.length;
      if (pos.offset >= s.index && pos.offset <= sEnd) {
        start = s.index;
        end = sEnd;
        if (pos.offset < sEnd) break; // boundary clicks prefer the following word
      }
    }
    deps.setSelection({
      anchor: { blockId: pos.blockId, offset: start },
      focus: { blockId: pos.blockId, offset: end },
    });
  };

  const onMouseDown = (ev: MouseEvent): void => {
    if (ev.button !== 0) return;
    deps.focusProxy();

    // An interactive embedder decoration (a badge/box/… with an onClick) is the
    // topmost overlay — a click on one is a UI action, not caret placement.
    if (deps.onDecorationClick(ev.clientX, ev.clientY)) {
      ev.preventDefault();
      return;
    }

    // ---- story-mode routing (Word semantics) ------------------------------
    const pt = deps.clientToPage(ev.clientX, ev.clientY);
    if (!pt) return;

    // Ctrl+click on a hyperlink opens it; on a TOC entry it jumps to the
    // heading (Word).
    if (ev.ctrlKey || ev.metaKey) {
      const href = linkAt(deps.getTree(), pt.pageIndex, pt.x, pt.y, scope());
      if (href) {
        ev.preventDefault();
        if (href.startsWith("#")) {
          // In-document anchor (TOC entry, cross-reference): scroll to the
          // target instead of opening a tab. The clicked paragraph's text is
          // used to resolve the heading when bookmarks aren't modeled.
          const pos = hitTest(deps.getTree(), pt.pageIndex, pt.x, pt.y, scope());
          deps.onAnchorJump(href.slice(1), pos?.blockId ?? null);
        } else {
          window.open(href, "_blank", "noopener");
        }
        return;
      }
      const pos = hitTest(deps.getTree(), pt.pageIndex, pt.x, pt.y, scope());
      const target = pos
        ? paragraphs().find((p) => p.id === pos.blockId)?.style.tocEntry?.targetId
        : undefined;
      if (target) {
        ev.preventDefault();
        deps.jumpToBlock(target);
        return;
      }
    }

    const band = bandAtPoint(pt);
    const story = deps.getStory();

    // ---- object layer: body images, incl. ones bleeding into the margins ----
    if (!story) {
      // A resize handle runs its own pointer-capture loop; the synthetic mousedown
      // it also emits must not re-run selection/caret logic here.
      if ((ev.target as HTMLElement | null)?.classList?.contains("cw-obj-handle")) return;
      // Editing a shape text box: a click INSIDE the box moves the caret there (it
      // must not re-grab the shape as an object); a click OUTSIDE commits and falls
      // through to normal body handling (#219).
      const editingShape = deps.getActiveShapeText();
      let stayInBox = false;
      if (editingShape) {
        const r = deps.shapeTextRect(editingShape);
        stayInBox =
          !!r &&
          pt.pageIndex === r.pageIndex &&
          pt.x >= r.x && pt.x < r.x + r.width &&
          pt.y >= r.y && pt.y < r.y + r.height;
        if (!stayInBox) deps.exitShapeText(false);
      }
      // Try to grab an object ONLY for a click genuinely inside the page (a gray-
      // area click is clamped onto the edge and must never "hit" a background
      // there — it should deselect, like Word). A background/floating image that
      // overlaps the header/footer margin still selects, since that's inside the
      // page. A double-click in a band falls through to enter the band story.
      if (pt.inside && !stayInBox && !(band && ev.detail >= 2)) {
        // Column/row grips live only in the body content area, never the margins.
        // Columns are tested FIRST so a cell-corner pixel prefers the column grip.
        if (!band) {
          const colHit = hitTestColumnBoundary(deps.getTree(), pt.pageIndex, pt.x, pt.y);
          if (colHit) {
            ev.preventDefault();
            deps.setColumnGuide(null); // the live border itself shows the drag now
            deps.startColumnDrag(colHit, ev);
            return;
          }
          const rowHit = hitTestRowBoundary(deps.getTree(), pt.pageIndex, pt.x, pt.y);
          if (rowHit) {
            ev.preventDefault();
            deps.setRowGuide(null); // the live row reflow itself shows the drag now
            deps.startRowDrag(rowHit, ev);
            return;
          }
        }
        const objHit = hitTestSelectableObject(deps.getTree(), pt.pageIndex, pt.x, pt.y, scope());
        if (objHit) {
          ev.preventDefault();
          // A double-click on a shape carrying a text box enters it (caret inside),
          // instead of just (re)selecting the object — like Word (#219). A double-
          // click on a top-level text-LESS shape starts a body and enters it (#235).
          if (ev.detail >= 2) {
            if (deps.shapeHasText(objHit.blockId)) {
              const sc: GeoScope = { shape: objHit.blockId, pageIndex: pt.pageIndex };
              const inside = hitTest(deps.getTree(), pt.pageIndex, pt.x, pt.y, sc);
              if (inside) {
                deps.enterShapeText(objHit.blockId, inside);
                return;
              }
            } else if (deps.startShapeText(objHit.blockId)) {
              return; // created an empty body + dropped the caret inside
            }
          }
          deps.selectObject(objHit.blockId);
          // Anchored images/shapes can be dragged to reposition: arm a move here.
          const mv = movableObject(objHit.blockId);
          if (mv) {
            const baseX = mv.anchor.offsetXPx;
            const baseY = mv.anchor.offsetYPx;
            // Back out the anchor's reference origin from the placed rect so a
            // snapped offset lands the object on the page-relative grid lines.
            const placed = objectRect(deps.getTree(), objHit.blockId);
            const origin = placed ? { ox: placed.x - baseX, oy: placed.y - baseY } : { ox: 0, oy: 0 };
            objectDrag = {
              blockId: objHit.blockId,
              startX: pt.x,
              startY: pt.y,
              baseX,
              baseY,
              lastX: baseX,
              lastY: baseY,
              origin,
              moved: false,
            };
          }
          return;
        }
      }
      // Nothing grabbed the click (empty space, off-page, or a band region) → drop
      // any object selection. The body then falls through to caret placement; a
      // band double-click proceeds to enter the band story below. (When staying in
      // an active text box, keep the caret there — don't clear the object, there is
      // none, and don't disturb the editing session.)
      if (!stayInBox) deps.selectObject(null);
    }
    if (story) {
      if (band === story.band) {
        // Same story, possibly a different page — re-pin the scope there.
        if (pt.pageIndex !== story.pageIndex) {
          deps.setStory({ band: story.band, pageIndex: pt.pageIndex });
        }
      } else if (ev.detail >= 2) {
        deps.setStory(null); // double-click outside the band exits to the body
        if (band) return; // ...unless it hit the OTHER band; require a fresh click
      } else {
        return; // single clicks outside the band are ignored in story mode
      }
    } else if (band) {
      if (ev.detail < 2) return; // single click in a band does nothing (Word)
      deps.setStory({ band, pageIndex: pt.pageIndex });
    }

    const pos = posFromEvent(ev);
    if (!pos) return;
    // A plain single click ON an inline equation selects it (its one U+FFFC char),
    // so it highlights and can be edited/deleted like a small inline object. A
    // shift-click falls through to the normal range-extend so a selection can grow
    // across the equation.
    if (ev.detail === 1 && pt.inside && !ev.shiftKey) {
      const eq = inlineEquationAt(deps.getTree(), pt.pageIndex, pt.x, pt.y, scope());
      if (eq) {
        ev.preventDefault();
        deps.setSelection({ anchor: { blockId: eq.blockId, offset: eq.start }, focus: { blockId: eq.blockId, offset: eq.end } });
        return;
      }
    }
    // Content controls intercept single clicks (checkbox toggles consume the
    // press; dropdown/date open beside the normally-placed caret).
    if (ev.detail === 1 && deps.onSdtPress(pos)) {
      ev.preventDefault();
      return;
    }
    ev.preventDefault();
    if (ev.detail >= 3) {
      deps.setSelection({
        anchor: { blockId: pos.blockId, offset: 0 },
        focus: { blockId: pos.blockId, offset: textOf(pos.blockId).length },
      });
      return;
    }
    if (ev.detail === 2) {
      selectWordAt(pos);
      return;
    }
    // Capture a potential cross-cell drag origin (body tables only). Dragging
    // into another cell of the same table switches to rectangular cell selection.
    const startCell = !story && !band ? hitTestCell(deps.getTree(), pt.pageIndex, pt.x, pt.y) : null;
    drag = {
      cellTable: startCell?.tableId ?? null,
      cellAnchor: startCell ? { row: startCell.row, col: startCell.col } : null,
    };
    deps.setCellSelection(null); // a fresh press clears any prior cell selection
    applyMove(pos, ev.shiftKey);
  };

  const stopAutoScroll = (): void => {
    autoScrollDir = 0;
    if (autoScrollRaf !== null) {
      cancelAnimationFrame(autoScrollRaf);
      autoScrollRaf = null;
    }
  };

  const autoScrollTick = (): void => {
    autoScrollRaf = null;
    if (!drag || autoScrollDir === 0) return;
    container.scrollTop += autoScrollDir;
    autoScrollRaf = requestAnimationFrame(autoScrollTick);
  };

  const onMouseMove = (ev: MouseEvent): void => {
    // Anchored-image drag takes priority over (and is exclusive with) text drag.
    if (objectDrag) {
      const pt = deps.clientToPage(ev.clientX, ev.clientY);
      if (!pt) return;
      const dx = pt.x - objectDrag.startX;
      const dy = pt.y - objectDrag.startY;
      if (!objectDrag.moved && Math.hypot(dx, dy) < 3) return; // a plain click stays a select
      objectDrag.moved = true;
      let nx = objectDrag.baseX + dx;
      let ny = objectDrag.baseY + dy;
      // Snap the ABSOLUTE position to the grid (then back to an offset), so the
      // object lands on the drawn lines whatever its relFrom* origin is.
      const g = deps.getGridSpacing();
      if (deps.isSnapToGrid() && g > 0) {
        const { ox, oy } = objectDrag.origin;
        nx = Math.round((ox + nx) / g) * g - ox;
        ny = Math.round((oy + ny) / g) * g - oy;
      }
      objectDrag.lastX = Math.round(nx);
      objectDrag.lastY = Math.round(ny);
      deps.applyObjectMove(objectDrag.blockId, objectDrag.lastX, objectDrag.lastY, true);
      return;
    }
    if (!drag) return;
    const rect = container.getBoundingClientRect();
    const margin = 28;
    autoScrollDir =
      ev.clientY < rect.top + margin
        ? -Math.ceil((rect.top + margin - ev.clientY) / 3)
        : ev.clientY > rect.bottom - margin
          ? Math.ceil((ev.clientY - (rect.bottom - margin)) / 3)
          : 0;
    if (autoScrollDir !== 0 && autoScrollRaf === null) autoScrollTick();
    // Cross-cell drag: once the pointer enters a different cell of the same
    // table, paint a rectangular cell selection instead of a text selection.
    if (drag.cellTable && drag.cellAnchor) {
      const pt = deps.clientToPage(ev.clientX, ev.clientY);
      const cur = pt ? hitTestCell(deps.getTree(), pt.pageIndex, pt.x, pt.y) : null;
      if (cur && cur.tableId === drag.cellTable) {
        const crossed = cur.row !== drag.cellAnchor.row || cur.col !== drag.cellAnchor.col;
        if (crossed) {
          deps.setCellSelection({ tableId: drag.cellTable, anchor: drag.cellAnchor, focus: { row: cur.row, col: cur.col } });
          return; // suppress text selection while spanning cells
        }
      }
    }
    const pos = posFromEvent(ev);
    if (pos) applyMove(pos, true);
  };

  const onMouseUp = (): void => {
    if (objectDrag) {
      if (objectDrag.moved) {
        // Revert the transient preview to the grab offset, then commit the whole
        // drag as ONE undoable op (mirrors the resize commit).
        deps.applyObjectMove(objectDrag.blockId, objectDrag.baseX, objectDrag.baseY, true);
        deps.applyObjectMove(objectDrag.blockId, objectDrag.lastX, objectDrag.lastY, false);
      }
      objectDrag = null;
    }
    drag = null;
    stopAutoScroll();
  };

  // Touch/pen: open the soft keyboard by focusing the IME proxy inside a trusted
  // gesture. The synthesized mousedown after a tap is NOT a trusted activation,
  // so focusing only there leaves the keyboard shut on iOS. We focus on pointerUP
  // (also the most reliable moment on iOS) and ONLY for a clean single-finger tap
  // — never when a second finger joined (pinch) or the finger dragged (scroll),
  // so pinch-zoom and scroll don't pop the keyboard. Caret placement still rides
  // the synthesized mouse path (onMouseDown); this only routes focus.
  let touchCount = 0;
  let touchTap: { x: number; y: number; multi: boolean } | null = null;
  const onTouchDown = (ev: PointerEvent): void => {
    if (ev.pointerType === "mouse") return;
    touchCount++;
    if (touchCount === 1) touchTap = { x: ev.clientX, y: ev.clientY, multi: false };
    else if (touchTap) touchTap.multi = true; // a 2nd finger → pinch, not a tap
  };
  const onTouchUp = (ev: PointerEvent): void => {
    if (ev.pointerType === "mouse") return;
    touchCount = Math.max(0, touchCount - 1);
    if (touchCount !== 0) return; // wait until every finger lifts
    const tap = touchTap;
    touchTap = null;
    if (!tap || tap.multi) return; // was a pinch / multi-touch
    if (Math.hypot(ev.clientX - tap.x, ev.clientY - tap.y) > 10) return; // a drag/scroll
    deps.focusProxy();
  };
  const onTouchCancel = (ev: PointerEvent): void => {
    if (ev.pointerType === "mouse") return;
    touchCount = Math.max(0, touchCount - 1);
    if (touchCount === 0) touchTap = null;
  };

  /** Is the point over a TOC-entry paragraph (Ctrl-clickable to its heading)? */
  const isTocEntryAt = (pt: { pageIndex: number; x: number; y: number }): boolean => {
    const pos = hitTest(deps.getTree(), pt.pageIndex, pt.x, pt.y, scope());
    return !!(pos && blockById(deps.getDoc(), pos.blockId)?.style.tocEntry);
  };

  // Pointer left the editor: drop any column/row-resize guide so it doesn't linger.
  const onHoverLeave = (): void => {
    hoverEvent = null; // a queued rAF must not re-apply hover state after leave
    if (!drag) {
      deps.setColumnGuide(null);
      deps.setRowGuide(null);
    }
  };

  // Hover affordances: col-resize over column grips, row-resize over row grips,
  // pointer + tooltip over links. mousemove can fire several times per frame
  // (high-polling mice), and each pass runs a chain of geometry probes — coalesce
  // to at most one probe pass per animation frame, on the latest pointer position.
  let hoverRaf: number | null = null;
  let hoverEvent: MouseEvent | null = null;
  const onHoverMove = (ev: MouseEvent): void => {
    hoverEvent = ev;
    if (hoverRaf !== null) return;
    hoverRaf = requestAnimationFrame(() => {
      hoverRaf = null;
      if (hoverEvent) processHoverMove(hoverEvent);
    });
  };
  const processHoverMove = (ev: MouseEvent): void => {
    if (drag) return;
    if (container.dataset["painter"]) {
      container.style.cursor = "copy"; // format painter armed
      return;
    }
    // An interactive decoration is the topmost overlay — a pointer cursor over one.
    if (deps.decorationAt(ev.clientX, ev.clientY) != null) {
      deps.setColumnGuide(null);
      deps.setRowGuide(null);
      container.style.cursor = "pointer";
      return;
    }
    const pt = deps.clientToPage(ev.clientX, ev.clientY);
    const inBody = !!pt && !deps.getStory();
    const colHit = inBody ? hitTestColumnBoundary(deps.getTree(), pt!.pageIndex, pt!.x, pt!.y) : null;
    // Columns win at a cell corner (tested first), so only probe rows clear of one.
    const rowHit = inBody && !colHit ? hitTestRowBoundary(deps.getTree(), pt!.pageIndex, pt!.x, pt!.y) : null;
    const hit = colHit || rowHit;
    const href = !hit && pt ? linkAt(deps.getTree(), pt.pageIndex, pt.x, pt.y, scope()) : null;
    // A TOC entry is Ctrl-clickable (jumps to its heading via tocEntry.targetId)
    // even with no run link — match the cursor to that click affordance, so
    // regenerated/link-free entries still read as clickable.
    // Paint a highlight on the boundary the pointer can grab (cleared otherwise).
    deps.setColumnGuide(
      colHit ? { pageIndex: colHit.pageIndex, x: colHit.x, y: colHit.tableY, height: colHit.tableHeight } : null,
    );
    deps.setRowGuide(
      rowHit ? { pageIndex: rowHit.pageIndex, y: rowHit.y, x: rowHit.tableX, width: rowHit.tableWidth } : null,
    );
    const overTocEntry = !hit && !href && pt ? isTocEntryAt(pt) : false;
    // Over an image, swap the I-beam for an object cursor: a move cursor on a
    // draggable anchored image, an arrow on an in-flow one.
    let objCursor: string | null = null;
    if (!hit && !href && !overTocEntry && pt?.inside && !deps.getStory()) {
      const oh = hitTestSelectableObject(deps.getTree(), pt.pageIndex, pt.x, pt.y, scope());
      if (oh) objCursor = movableObject(oh.blockId) ? "move" : "default";
    }
    container.style.cursor = colHit
      ? "col-resize"
      : rowHit
        ? "row-resize"
        : href || overTocEntry
          ? "pointer"
          : objCursor ?? "text";
    if ((container.title || "") !== (href ?? "")) container.title = href ?? "";
  };

  // ---- keyboard ------------------------------------------------------------

  const onKeyDown = (ev: KeyboardEvent): void => {
    const sel = deps.getSelection();
    const tree = deps.getTree();
    const ctrl = ev.ctrlKey || ev.metaKey;
    const sc = scope();

    if (ev.key === "Escape") {
      // Editing a text box: Escape pops back to selecting the shape object (#219).
      if (deps.getActiveShapeText()) {
        deps.exitShapeText(true);
        ev.preventDefault();
        return;
      }
      if (deps.hasSelectedObject()) {
        deps.selectObject(null);
        ev.preventDefault();
        return;
      }
      if (deps.getStory()) {
        deps.setStory(null); // exit story mode, restoring the body selection
        ev.preventDefault();
        return;
      }
    }
    if ((ev.key === "Delete" || ev.key === "Backspace") && deps.hasSelectedObject()) {
      deps.deleteSelectedObject();
      ev.preventDefault();
      return;
    }
    // F6 (Shift+F6 backward): step the keyboard focus into the drawing-object layer
    // and cycle through every selectable object — the "reach the objects" gesture a
    // canvas gives no other keyboard path to (E3). Works from a text caret too.
    if (ev.key === "F6" && !ctrl && !ev.altKey && !deps.getActiveShapeText()) {
      if (deps.cycleObjectFocus(ev.shiftKey)) {
        ev.preventDefault();
        return;
      }
    }
    // Enter on a selected shape carrying a text box enters it (caret inside) — the
    // keyboard mirror of the double-click-to-edit gesture (#219).
    if (ev.key === "Enter" && !ctrl && !ev.shiftKey && deps.hasSelectedObject()) {
      if (deps.enterSelectedShapeText()) {
        ev.preventDefault();
        return;
      }
    }
    // Once in the object layer, Tab / Shift+Tab walk to the next / previous object
    // (so a keyboard user can move between drawings, then nudge/rotate/delete them —
    // E3). Plain Tab elsewhere keeps its cell-navigation / default behaviour.
    if (ev.key === "Tab" && !ctrl && deps.hasSelectedObject()) {
      if (deps.cycleObjectFocus(ev.shiftKey)) {
        ev.preventDefault();
        return;
      }
    }
    if (ev.key === "Tab" && !ctrl) {
      if (deps.onTab(ev.shiftKey)) {
        ev.preventDefault();
        return;
      }
    }
    if (ctrl && ev.key.toLowerCase() === "a") {
      const edges = documentEdges(tree, sc); // scoped: selects the band story
      if (edges) deps.setSelection({ anchor: edges.start, focus: edges.end });
      ev.preventDefault();
      return;
    }
    // Arrow-nudge a selected object (E2): a selected object clears the text
    // selection (sel is null), so this must run BEFORE the text-navigation guard.
    // 1px per press, 10px with Shift; the wiring clamps to the page and only an
    // anchored shape actually moves (returns true → the arrow is consumed).
    if (deps.hasSelectedObject() && !ctrl && ARROW_DELTA[ev.key]) {
      const [ux, uy] = ARROW_DELTA[ev.key]!;
      const step = ev.shiftKey ? 10 : 1;
      if (deps.nudgeSelectedObject(ux * step, uy * step)) {
        ev.preventDefault();
        return;
      }
    }
    if (!sel) return;

    const collapseTo = (which: "min" | "max"): DocPosition => {
      const cmp = comparePositions(tree, sel.anchor, sel.focus, sc);
      const [min, max] = cmp <= 0 ? [sel.anchor, sel.focus] : [sel.focus, sel.anchor];
      return which === "min" ? min : max;
    };

    switch (ev.key) {
      case "ArrowLeft":
      case "ArrowRight": {
        const dir: -1 | 1 = ev.key === "ArrowLeft" ? -1 : 1;
        // A plain arrow on a non-collapsed selection collapses to its edge
        // without moving (Word/native behavior); otherwise step from focus.
        let target: DocPosition;
        if (!ev.shiftKey && !isCollapsed(sel) && !ctrl) {
          // Visual edge: in an RTL paragraph "min" (logical) is the right edge.
          const rtl = paragraphs().find((b) => b.id === sel.focus.blockId)?.style.direction === "rtl";
          const leftEdge = rtl ? "max" : "min";
          target = collapseTo(dir === -1 ? leftEdge : leftEdge === "min" ? "max" : "min");
        } else {
          target = moveHorizontalVisual(sel.focus, dir, ctrl);
        }
        applyMove(target, ev.shiftKey);
        ev.preventDefault();
        return;
      }
      case "ArrowUp":
      case "ArrowDown": {
        const goalX = sel.goalX ?? caretRect(tree, sel.focus, sc)?.x ?? 0;
        const next = positionOnAdjacentLine(
          tree,
          sel.focus,
          ev.key === "ArrowUp" ? "up" : "down",
          goalX,
          sc,
        );
        if (next) applyMove(next, ev.shiftKey, goalX);
        ev.preventDefault();
        return;
      }
      case "Home":
      case "End": {
        if (ctrl) {
          const edges = documentEdges(tree, sc);
          if (edges) applyMove(ev.key === "Home" ? edges.start : edges.end, ev.shiftKey);
        } else {
          const edges = lineEdges(tree, sel.focus, sc);
          if (edges) applyMove(ev.key === "Home" ? edges.home : edges.end, ev.shiftKey);
        }
        ev.preventDefault();
        return;
      }
    }
  };

  // ---- copy / cut (text/html + text/plain flavors) -------------------------

  const writeSelectionToClipboard = (ev: ClipboardEvent): boolean => {
    // A rectangular cell selection serializes the whole grid rectangle as an
    // HTML table + tab-separated text, ahead of any (stale) text selection.
    const cell = deps.getCellSelection();
    if (cell) {
      const found = findTableById(deps.getDoc(), cell.tableId);
      if (found) {
        const { html, text } = tableRectToClipboard(found.table, {
          r0: cell.anchor.row,
          c0: cell.anchor.col,
          r1: cell.focus.row,
          c1: cell.focus.col,
        });
        ev.clipboardData?.setData("text/html", html);
        ev.clipboardData?.setData("text/plain", text);
        ev.preventDefault();
        return true;
      }
    }
    const sel = deps.getSelection();
    if (!sel || isCollapsed(sel)) return false;
    const tree = deps.getTree();
    const cmp = comparePositions(tree, sel.anchor, sel.focus, scope());
    const [from, to] = cmp <= 0 ? [sel.anchor, sel.focus] : [sel.focus, sel.anchor];
    const fragment = extractFragment(paragraphs(), from, to);
    ev.clipboardData?.setData("text/html", fragmentToHtml(fragment));
    ev.clipboardData?.setData("text/plain", fragmentToPlainText(fragment));
    ev.preventDefault();
    return true;
  };

  const onCopy = (ev: ClipboardEvent): void => {
    writeSelectionToClipboard(ev);
  };

  const onCut = (ev: ClipboardEvent): void => {
    if (writeSelectionToClipboard(ev)) deps.onDeleteSelection();
  };

  container.addEventListener("mousedown", onMouseDown);
  container.addEventListener("pointerdown", onTouchDown);
  container.addEventListener("pointerup", onTouchUp);
  container.addEventListener("pointercancel", onTouchCancel);
  window.addEventListener("mousemove", onMouseMove);
  window.addEventListener("mouseup", onMouseUp);
  container.addEventListener("mousemove", onHoverMove);
  container.addEventListener("mouseleave", onHoverLeave);
  container.addEventListener("keydown", onKeyDown);
  container.addEventListener("copy", onCopy);
  container.addEventListener("cut", onCut);

  return {
    destroy(): void {
      stopAutoScroll();
      if (hoverRaf !== null) {
        cancelAnimationFrame(hoverRaf);
        hoverRaf = null;
      }
      container.removeEventListener("mousedown", onMouseDown);
      container.removeEventListener("pointerdown", onTouchDown);
      container.removeEventListener("pointerup", onTouchUp);
      container.removeEventListener("pointercancel", onTouchCancel);
      window.removeEventListener("mousemove", onMouseMove);
      window.removeEventListener("mouseup", onMouseUp);
      container.removeEventListener("mousemove", onHoverMove);
      container.removeEventListener("mouseleave", onHoverLeave);
      container.removeEventListener("keydown", onKeyDown);
      container.removeEventListener("copy", onCopy);
      container.removeEventListener("cut", onCut);
    },
  };
}
