// Mouse/keyboard selection over the canvas, driven entirely by layout hit-testing.
//
// - mousedown -> hitTest -> anchor; drag extends focus; auto-scroll near edges.
// - double-click: word select via Intl.Segmenter (word granularity) — the same
//   segmentation pretext uses, so selections align with break opportunities.
// - triple-click: paragraph select.
// - Arrows move by grapheme cluster, Ctrl+arrow by word, Home/End by line
//   (a layout query, not a model one), Up/Down preserve goalX, shift extends,
//   Ctrl+A selects all, Ctrl+C copies plain text (HTML flavor lands in milestone 5).

import type { Document, Paragraph } from "@cw/shared";
import { BAND_CONTAINERS } from "@cw/shared";
import type { DocPosition, DocSelection } from "@cw/shared";
import { isCollapsed } from "@cw/shared";
import {
  words,
  bandParagraphs,
  paragraphsOf,
  prevGrapheme,
  nextGrapheme,
  prevWordStart,
  nextWordEnd,
  type BandName,
} from "@cw/shared";
import type { LayoutTree } from "../layout/layoutTree";
import {
  comparePositions,
  documentEdges,
  hitTest,
  hitTestCell,
  hitTestObject,
  hitTestColumnBoundary,
  caretRect,
  lineEdges,
  linkAt,
  positionOnAdjacentLine,
  type ColumnBoundaryHit,
  type GeoScope,
} from "../layout/geometry";
import type { PagePoint } from "../paint/renderer";
import type { CellSelection } from "../editor/state";
import { extractFragment, fragmentToHtml, fragmentToPlainText } from "./clipboard";

export interface SelectionControllerDeps {
  container: HTMLElement;
  getTree(): LayoutTree;
  getDoc(): Document;
  getSelection(): DocSelection | null;
  setSelection(sel: DocSelection | null): void;
  /** Set (or clear) the rectangular table-cell selection during a cross-cell drag. */
  setCellSelection(sel: CellSelection | null): void;
  clientToPage(clientX: number, clientY: number): PagePoint | null;
  /** Route focus to the IME proxy so typing works right after a click. */
  focusProxy(): void;
  /** Delete the current selection (Ctrl+X after the copy half). */
  onDeleteSelection(): void;
  /** Story-edit scope (header/footer mode). Null = editing the body. */
  getStory(): GeoScope | null;
  /** Enter (band+page) or exit (null) story mode; triggers relayout + dimming. */
  setStory(scope: GeoScope | null): void;
  /** Select an image object (null clears). */
  selectObject(blockId: string | null): void;
  /** True when an object is selected (Delete/Backspace/Escape route to it). */
  hasSelectedObject(): boolean;
  deleteSelectedObject(): void;
  /** Begin a table column-boundary drag (wiring owns the transient/commit loop). */
  startColumnDrag(hit: ColumnBoundaryHit, ev: MouseEvent): void;
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
}

export interface SelectionController {
  destroy(): void;
}

export function createSelectionController(deps: SelectionControllerDeps): SelectionController {
  const { container } = deps;
  container.tabIndex = 0;
  container.style.outline = "none";
  container.style.cursor = "text";

  // ---- model text helpers ----------------------------------------------

  const scope = (): GeoScope | undefined => deps.getStory() ?? undefined;

  // Navigation-order paragraphs for the ACTIVE story: while story-editing,
  // the variant CONTAINER that renders on the story's page (first/even bands
  // override the default); otherwise the body (including table cells).
  const paragraphs = (): Paragraph[] => {
    const story = deps.getStory();
    if (story) {
      const pg = deps.getTree().pages[story.pageIndex];
      const source =
        (story.band === "header" ? pg?.headerSource : pg?.footerSource) ?? story.band;
      return bandParagraphs(deps.getDoc(), source);
    }
    return paragraphsOf(deps.getDoc()).filter((p) => !isBandParagraph(p.id));
  };

  const isBandParagraph = (blockId: string): boolean => {
    const doc = deps.getDoc();
    return BAND_CONTAINERS.some((band) => bandParagraphs(doc, band).some((p) => p.id === blockId));
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
    if (dir === -1) {
      if (pos.offset > 0) {
        const offset = byWord ? prevWordStart(text, pos.offset) : prevGrapheme(text, pos.offset);
        return { blockId: pos.blockId, offset };
      }
      const prev = blocks[bi - 1];
      return prev ? { blockId: prev.id, offset: textOf(prev.id).length } : pos;
    }
    if (pos.offset < text.length) {
      const offset = byWord ? nextWordEnd(text, pos.offset) : nextGrapheme(text, pos.offset);
      return { blockId: pos.blockId, offset };
    }
    const next = blocks[bi + 1];
    return next ? { blockId: next.id, offset: 0 } : pos;
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

  let dragging = false;
  let autoScrollDir = 0;
  let autoScrollRaf: number | null = null;
  // Cross-cell drag: the table + anchor cell captured on mousedown, and whether
  // the drag has crossed into another cell (switching from text to cell selection).
  let cellDragTable: string | null = null;
  let cellDragAnchor: { row: number; col: number } | null = null;

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

    // ---- object layer: column grips and image selection (body mode only) ----
    if (!story && !band) {
      const colHit = hitTestColumnBoundary(deps.getTree(), pt.pageIndex, pt.x, pt.y);
      if (colHit) {
        ev.preventDefault();
        deps.startColumnDrag(colHit, ev);
        return;
      }
      const objHit = hitTestObject(deps.getTree(), pt.pageIndex, pt.x, pt.y);
      if (objHit) {
        ev.preventDefault();
        deps.selectObject(objHit.blockId);
        return;
      }
      deps.selectObject(null); // clicking text/empty space deselects objects
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
    cellDragTable = startCell?.tableId ?? null;
    cellDragAnchor = startCell ? { row: startCell.row, col: startCell.col } : null;
    deps.setCellSelection(null); // a fresh press clears any prior cell selection
    dragging = true;
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
    if (!dragging || autoScrollDir === 0) return;
    container.scrollTop += autoScrollDir;
    autoScrollRaf = requestAnimationFrame(autoScrollTick);
  };

  const onMouseMove = (ev: MouseEvent): void => {
    if (!dragging) return;
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
    if (cellDragTable && cellDragAnchor) {
      const pt = deps.clientToPage(ev.clientX, ev.clientY);
      const cur = pt ? hitTestCell(deps.getTree(), pt.pageIndex, pt.x, pt.y) : null;
      if (cur && cur.tableId === cellDragTable) {
        const crossed = cur.row !== cellDragAnchor.row || cur.col !== cellDragAnchor.col;
        if (crossed) {
          deps.setCellSelection({ tableId: cellDragTable, anchor: cellDragAnchor, focus: { row: cur.row, col: cur.col } });
          return; // suppress text selection while spanning cells
        }
      }
    }
    const pos = posFromEvent(ev);
    if (pos) applyMove(pos, true);
  };

  const onMouseUp = (): void => {
    dragging = false;
    cellDragTable = null;
    cellDragAnchor = null;
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

  // Hover affordances: col-resize over column grips, pointer + tooltip over links.
  const onHoverMove = (ev: MouseEvent): void => {
    if (dragging) return;
    if (container.dataset["painter"]) {
      container.style.cursor = "copy"; // format painter armed
      return;
    }
    const pt = deps.clientToPage(ev.clientX, ev.clientY);
    const hit =
      pt && !deps.getStory()
        ? hitTestColumnBoundary(deps.getTree(), pt.pageIndex, pt.x, pt.y)
        : null;
    const href = !hit && pt ? linkAt(deps.getTree(), pt.pageIndex, pt.x, pt.y, scope()) : null;
    container.style.cursor = hit ? "col-resize" : href ? "pointer" : "text";
    if ((container.title || "") !== (href ?? "")) container.title = href ?? "";
  };

  // ---- keyboard ------------------------------------------------------------

  const onKeyDown = (ev: KeyboardEvent): void => {
    const sel = deps.getSelection();
    const tree = deps.getTree();
    const ctrl = ev.ctrlKey || ev.metaKey;
    const sc = scope();

    if (ev.key === "Escape") {
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
        const target =
          !ev.shiftKey && !isCollapsed(sel) && !ctrl
            ? collapseTo(dir === -1 ? "min" : "max")
            : moveHorizontal(sel.focus, dir, ctrl);
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
  container.addEventListener("keydown", onKeyDown);
  container.addEventListener("copy", onCopy);
  container.addEventListener("cut", onCut);

  return {
    destroy(): void {
      stopAutoScroll();
      container.removeEventListener("mousedown", onMouseDown);
      container.removeEventListener("pointerdown", onTouchDown);
      container.removeEventListener("pointerup", onTouchUp);
      container.removeEventListener("pointercancel", onTouchCancel);
      window.removeEventListener("mousemove", onMouseMove);
      window.removeEventListener("mouseup", onMouseUp);
      container.removeEventListener("mousemove", onHoverMove);
      container.removeEventListener("keydown", onKeyDown);
      container.removeEventListener("copy", onCopy);
      container.removeEventListener("cut", onCut);
    },
  };
}
