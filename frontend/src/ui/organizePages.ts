// "Organize Pages" — a visual page-reorder overlay (PDF-style page sorter) for a
// docx. Each rendered page is a live thumbnail; pages that belong to the same
// PAGE GROUP (a run of pages between hard breaks — see ../layout/pageGroups) are
// bracketed together and move as one unit, so reordering only re-sequences whole
// blocks and never splits a paragraph or corrupts a section.
//
// Interactions: drag a group to a new slot (drop indicator snaps to group gaps),
// click × to delete a group, double-click a page to jump the editor there. Every
// edit dispatches an undoable command against the live document and the grid
// rebuilds from the new layout, so the overlay always shows committed state.

import { createDialogShell } from "./dialogShell";
import { createPaintLayer, type PaintScheduler } from "../paint/renderer";
import { computePageGroups, type PageGroup } from "../layout/pageGroups";
import { deletePageGroupCmd, pinPageBreakBeforeCmd, reorderPageGroupCmd } from "../editor/commands";
import type { Command } from "../editor/state";
import type { LayoutTree } from "../layout/layoutTree";
import type { ResolvedTheme } from "../config";
import type { CustomFontRegistry } from "../fonts/customRegistry";
import type { Document } from "@cw/shared";

/** The slice of the editor this overlay needs (structural — avoids importing the
 *  whole Editor type and its cycle). */
export interface OrganizePagesEditor {
  getDocument(): Document;
  getLayoutTree(): LayoutTree;
  dispatch(cmd: Command): void;
  revealBlock(blockId: string): void;
}

export interface OrganizePagesOptions {
  editor: OrganizePagesEditor;
  /** This editor instance's font registry, so thumbnails render with its fonts. */
  fontRegistry?: CustomFontRegistry;
  /** This editor instance's theme (page colors etc.). */
  theme?: ResolvedTheme;
  onClose?: () => void;
}

export interface OrganizePagesHandle {
  close(): void;
}

/** Thumbnail width in CSS px (height derives from each page's aspect ratio). */
const THUMB_W = 120;
/** Pointer travel (px) before a press becomes a drag (vs a click/double-click). */
const DRAG_THRESHOLD = 5;

const CSS = `
.cw-orgpg-backdrop{position:fixed;inset:0;z-index:2147483000;pointer-events:none;}
.cw-orgpg-modal{position:fixed;pointer-events:auto;
  width:min(92vw,880px);max-height:86vh;display:flex;flex-direction:column;background:#fff;color:#201f1e;
  border:1px solid #c8c6c4;border-radius:8px;box-shadow:0 12px 40px rgba(0,0,0,.28);font:13px 'Segoe UI',Roboto,sans-serif;}
.cw-orgpg-head{display:flex;align-items:center;gap:8px;padding:10px 12px;border-bottom:1px solid #edebe9;cursor:move;}
.cw-orgpg-head h2{margin:0;font-size:15px;font-weight:600;flex:1;}
.cw-orgpg-x{border:0;background:transparent;font-size:20px;line-height:1;cursor:pointer;color:#605e5c;padding:2px 6px;border-radius:4px;}
.cw-orgpg-x:hover{background:#f3f2f1;color:#201f1e;}
.cw-orgpg-hint{padding:8px 14px 0;color:#605e5c;font-size:12px;flex:0 0 auto;}
.cw-orgpg-body{padding:12px 14px 20px;overflow:auto;flex:1 1 auto;min-height:0;}
.cw-orgpg-grid{position:relative;display:flex;flex-wrap:wrap;gap:16px;align-content:flex-start;}
.cw-orgpg-group{position:relative;display:flex;flex-wrap:wrap;gap:6px;max-width:100%;padding:8px;border:1px solid #e1dfdd;
  border-radius:8px;background:#faf9f8;cursor:grab;touch-action:none;transition:box-shadow .12s,border-color .12s,opacity .12s;}
.cw-orgpg-group:hover{border-color:#c7c5c3;box-shadow:0 2px 8px rgba(0,0,0,.08);}
.cw-orgpg-group.linked{border-color:#c7b8e6;background:#f7f3fc;}
.cw-orgpg-group.dragging{opacity:.45;cursor:grabbing;box-shadow:0 6px 18px rgba(0,0,0,.22);}
.cw-orgpg-tile{display:flex;flex-direction:column;align-items:center;gap:4px;}
.cw-orgpg-thumb{width:${THUMB_W}px;background:#fff;border:1px solid #d0cfce;border-radius:2px;overflow:hidden;
  display:flex;align-items:flex-start;justify-content:center;}
.cw-orgpg-thumb.ph{align-items:center;justify-content:center;color:#a19f9d;font-size:12px;}
.cw-orgpg-num{font-size:11px;color:#605e5c;}
.cw-orgpg-tile.cur .cw-orgpg-thumb{border-color:#6b3fa0;box-shadow:0 0 0 2px rgba(107,63,160,.25);}
.cw-orgpg-del{position:absolute;top:-9px;right:-9px;width:20px;height:20px;border-radius:50%;border:1px solid #d0cfce;
  background:#fff;color:#a4262c;font-size:13px;line-height:1;cursor:pointer;display:none;align-items:center;justify-content:center;
  box-shadow:0 1px 3px rgba(0,0,0,.2);z-index:2;}
.cw-orgpg-group:hover .cw-orgpg-del{display:flex;}
.cw-orgpg-del:hover{background:#fde7e9;}
.cw-orgpg-badge{position:absolute;left:8px;bottom:-9px;background:#6b3fa0;color:#fff;font-size:10px;font-weight:600;
  padding:1px 6px;border-radius:8px;white-space:nowrap;}
.cw-orgpg-marker{position:absolute;width:3px;background:#6b3fa0;border-radius:2px;pointer-events:none;display:none;z-index:5;}
.cw-orgpg-foot{padding:10px 14px;border-top:1px solid #edebe9;display:flex;justify-content:flex-end;gap:8px;}
.cw-orgpg-btn{border:1px solid #c8c6c4;background:#fff;border-radius:4px;padding:6px 16px;font-size:13px;cursor:pointer;}
.cw-orgpg-btn.primary{background:#6b3fa0;border-color:#6b3fa0;color:#fff;}
.cw-orgpg-empty{padding:24px;text-align:center;color:#605e5c;}
`;

export function showOrganizePages(opts: OrganizePagesOptions): OrganizePagesHandle {
  const { editor } = opts;
  const shell = createDialogShell({
    prefix: "cw-orgpg",
    cssId: "cw-orgpg-css",
    css: CSS,
    title: "Organize Pages",
    extraNoDrag: ".cw-orgpg-body",
    onClose: () => {
      teardownThumbs();
      opts.onClose?.();
    },
  });

  const hint = document.createElement("div");
  hint.className = "cw-orgpg-hint";
  hint.textContent = "Drag a page to reorder. Linked pages (no break between them) move together. Double-click a page to jump to it.";
  shell.modal.insertBefore(hint, shell.body);

  const grid = document.createElement("div");
  grid.className = "cw-orgpg-grid";
  shell.body.appendChild(grid);

  const marker = document.createElement("div");
  marker.className = "cw-orgpg-marker";
  grid.appendChild(marker);

  const doneBtn = document.createElement("button");
  doneBtn.className = "cw-orgpg-btn primary";
  doneBtn.textContent = "Done";
  doneBtn.addEventListener("click", () => shell.close());
  shell.foot.appendChild(doneBtn);

  // ---- live thumbnail rendering (one paint layer per page) ----------------
  // Rendered eagerly on (re)build; capped so a pathologically long document
  // can't spawn thousands of canvases (tiles past the cap still show their
  // number and remain reorderable).
  const MAX_THUMBS = 150;
  const painters: PaintScheduler[] = [];
  let renderBudget = 0;
  /** The movable units of the last rebuild — drag/delete read their block ranges. */
  let units: PageGroup[] = [];
  const teardownThumbs = (): void => {
    for (const p of painters) p.destroy();
    painters.length = 0;
  };

  const renderThumb = (host: HTMLElement, tree: LayoutTree, pageIndex: number): void => {
    const page = tree.pages[pageIndex];
    if (!page) return;
    const scale = THUMB_W / page.widthPx;
    // The paint layer keys its placeholder + zoom off page.index, so a single-page
    // tree must present the page AT index 0 (a shallow clone — the original is
    // shared with the live tree). zoomMin defaults to 0.25, which would clamp a
    // small thumbnail scale, so lower it for this layer.
    const single: LayoutTree = { pages: [{ ...page, index: 0 }], pageWidthPx: page.widthPx, pageHeightPx: page.heightPx, marginPx: page.marginPx };
    const painter = createPaintLayer(host, {
      chrome: false,
      zoomMin: Math.min(0.05, scale),
      ...(opts.theme ? { theme: opts.theme } : {}),
      ...(opts.fontRegistry ? { fontRegistry: opts.fontRegistry } : {}),
    });
    painter.setTree(single);
    painter.setZoom(scale);
    painters.push(painter);
  };

  // ---- layout-verified pagination repair ----------------------------------
  // A reorder can reshuffle the section-break chain (or strand a skipped hidden
  // first paragraph) so a group merges onto the previous page — it renders zero
  // pages. Pin an explicit break on each such group's first block until the
  // layout is clean. Surgical: only actually-merged groups are touched, so we
  // never add a blank page. Bounded so a group we can't repair (non-paragraph
  // first block) can't loop forever.
  const repairPagination = (): void => {
    for (let iter = 0; iter < 12; iter++) {
      const groups = computePageGroups(editor.getDocument(), editor.getLayoutTree());
      const merged = groups.find((g) => g.pageIndices.length === 0 && g.blockIds.length > 0);
      if (!merged) return;
      const before = editor.getDocument();
      editor.dispatch(pinPageBreakBeforeCmd(merged.blockIds[0]!));
      if (editor.getDocument() === before) return; // couldn't repair (non-paragraph first) — stop
    }
  };

  // ---- grid (re)build from the current document + layout ------------------
  const rebuild = (): void => {
    teardownThumbs();
    // Remove everything except the persistent marker.
    for (const el of [...grid.children]) if (el !== marker) el.remove();

    const doc = editor.getDocument();
    const tree = editor.getLayoutTree();
    const groups = computePageGroups(doc, tree);
    units = groups;

    if (groups.length === 0 || tree.pages.length === 0) {
      const empty = document.createElement("div");
      empty.className = "cw-orgpg-empty";
      empty.textContent = "This document has no pages to organize.";
      grid.appendChild(empty);
      return;
    }

    renderBudget = MAX_THUMBS;
    for (const g of groups) buildGroup(g, tree, groups.length);
  };

  const buildGroup = (g: PageGroup, tree: LayoutTree, groupCount: number): void => {
    const groupEl = document.createElement("div");
    groupEl.className = "cw-orgpg-group" + (g.pageIndices.length > 1 ? " linked" : "");
    groupEl.dataset["group"] = String(g.index);

    // Delete (whole group). Hidden when it is the only group — never empty the doc.
    if (groupCount > 1) {
      const del = document.createElement("button");
      del.className = "cw-orgpg-del";
      del.textContent = "×";
      del.title = g.pageIndices.length > 1 ? "Delete these pages" : "Delete this page";
      del.addEventListener("pointerdown", (e) => e.stopPropagation());
      del.addEventListener("click", (e) => {
        e.stopPropagation();
        editor.dispatch(deletePageGroupCmd(g.blockIds[0]!, g.blockIds[g.blockIds.length - 1]!));
        repairPagination();
        rebuild();
      });
      groupEl.appendChild(del);
    }

    const pages = g.pageIndices.length > 0 ? g.pageIndices : [-1];
    for (const pi of pages) {
      const tile = document.createElement("div");
      tile.className = "cw-orgpg-tile";
      const thumb = document.createElement("div");
      thumb.className = "cw-orgpg-thumb";
      const num = document.createElement("div");
      num.className = "cw-orgpg-num";
      if (pi >= 0) {
        const page = tree.pages[pi]!;
        thumb.style.height = `${Math.round(page.heightPx * (THUMB_W / page.widthPx))}px`;
        thumb.dataset["page"] = String(pi);
        tile.dataset["page"] = String(pi);
        const firstBlock = page.blocks[0]?.blockId;
        if (firstBlock) tile.dataset["block"] = firstBlock;
        num.textContent = String(page.number);
        if (renderBudget > 0) { renderThumb(thumb, tree, pi); renderBudget--; }
      } else {
        thumb.classList.add("ph");
        thumb.style.height = `${Math.round(THUMB_W * 1.294)}px`;
        thumb.textContent = "—";
        num.textContent = "—";
      }
      tile.append(thumb, num);
      groupEl.appendChild(tile);
    }

    if (g.pageIndices.length > 1) {
      const badge = document.createElement("div");
      badge.className = "cw-orgpg-badge";
      const first = tree.pages[g.pageIndices[0]!]!.number;
      const last = tree.pages[g.pageIndices[g.pageIndices.length - 1]!]!.number;
      badge.textContent = `Pages ${first}–${last}`;
      badge.title = "These pages have no break between them, so they move together.";
      groupEl.appendChild(badge);
    }

    grid.appendChild(groupEl);
  };

  // ---- double-click a page → jump the editor there (panel stays open) -----
  // The overlay is a non-blocking floating panel, so we scroll the canvas
  // behind it and leave it open — the user can keep reorganizing.
  grid.addEventListener("dblclick", (e) => {
    if (dragging) return;
    const tile = (e.target as HTMLElement).closest(".cw-orgpg-tile") as HTMLElement | null;
    const blockId = tile?.dataset["block"];
    if (blockId) editor.revealBlock(blockId);
  });

  // ---- drag-to-reorder (whole group is the unit) --------------------------
  let dragging: { groupIndex: number; el: HTMLElement; startX: number; startY: number; started: boolean } | null = null;
  let pendingSlot: number | null = null;

  const groupEls = (): HTMLElement[] => [...grid.querySelectorAll(".cw-orgpg-group")] as HTMLElement[];

  /** Nearest insertion slot (group-gap index in 0..count) for a pointer point. */
  const slotAt = (clientX: number, clientY: number): number => {
    const els = groupEls();
    let best = els.length;
    let bestD = Infinity;
    els.forEach((el, i) => {
      const r = el.getBoundingClientRect();
      const cx = r.left + r.width / 2;
      const cy = r.top + r.height / 2;
      const d = Math.hypot(clientX - cx, clientY - cy);
      if (d < bestD) {
        bestD = d;
        best = clientX < cx ? i : i + 1;
      }
    });
    return best;
  };

  const showMarker = (slot: number): void => {
    const els = groupEls();
    const gridRect = grid.getBoundingClientRect();
    let x: number;
    let top: number;
    let height: number;
    if (slot < els.length) {
      const r = els[slot]!.getBoundingClientRect();
      x = r.left - gridRect.left - 8;
      top = r.top - gridRect.top;
      height = r.height;
    } else {
      const r = els[els.length - 1]!.getBoundingClientRect();
      x = r.right - gridRect.left + 5;
      top = r.top - gridRect.top;
      height = r.height;
    }
    marker.style.left = `${x}px`;
    marker.style.top = `${top}px`;
    marker.style.height = `${height}px`;
    marker.style.display = "block";
  };

  const onPointerDown = (e: PointerEvent): void => {
    if (e.button !== 0) return;
    const groupEl = (e.target as HTMLElement).closest(".cw-orgpg-group") as HTMLElement | null;
    if (!groupEl) return;
    dragging = { groupIndex: Number(groupEl.dataset["group"]), el: groupEl, startX: e.clientX, startY: e.clientY, started: false };
    pendingSlot = null;
  };

  const onPointerMove = (e: PointerEvent): void => {
    if (!dragging) return;
    if (!dragging.started) {
      if (Math.hypot(e.clientX - dragging.startX, e.clientY - dragging.startY) < DRAG_THRESHOLD) return;
      dragging.started = true;
      dragging.el.classList.add("dragging");
      grid.style.userSelect = "none";
    }
    pendingSlot = slotAt(e.clientX, e.clientY);
    showMarker(pendingSlot);
  };

  const endDrag = (): void => {
    if (!dragging) return;
    const d = dragging;
    dragging = null;
    marker.style.display = "none";
    grid.style.userSelect = "";
    d.el.classList.remove("dragging");
    if (d.started && pendingSlot !== null) {
      const moved = units[d.groupIndex];
      if (moved) {
        const anchor = pendingSlot < units.length ? units[pendingSlot] : undefined;
        const anchorId = anchor ? anchor.blockIds[0]! : null;
        editor.dispatch(reorderPageGroupCmd(moved.blockIds[0]!, moved.blockIds[moved.blockIds.length - 1]!, anchorId));
        repairPagination();
        rebuild();
      }
    }
    pendingSlot = null;
  };

  grid.addEventListener("pointerdown", onPointerDown, { signal: shell.signal });
  window.addEventListener("pointermove", onPointerMove, { signal: shell.signal });
  window.addEventListener("pointerup", endDrag, { signal: shell.signal });

  rebuild();
  return { close: () => shell.close() };
}
