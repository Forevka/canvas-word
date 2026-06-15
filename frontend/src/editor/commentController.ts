// Floating comment affordance — the 💬 chip beside a ranged suggest-mode selection
// and the inline composer it expands into. Extracted from createEditor (index.ts)
// so the editor factory no longer owns this DOM + lifecycle. Pure presentation:
// the host injects live selection/doc/layout/mode getters and the addComment sink.

import type { CharStyle, Document, DocSelection, Fragment, UserInfo } from "@cw/shared";
import { blockById, colorForId, isCollapsed, styleAtRuns } from "@cw/shared";
import { attachMentionAutocomplete } from "../review/mentions";
import { selectionRects, type CaretRect, type GeoScope } from "../layout/geometry";
import type { LayoutTree } from "../layout/layoutTree";
import type { CellSelection, EditMode } from "./state";

export interface CommentControllerDeps {
  container: HTMLElement;
  caretToContainer: (caret: CaretRect) => { left: number; top: number } | null;
  getSelection: () => DocSelection | null;
  getCellSelection: () => CellSelection | null | undefined;
  getTree: () => LayoutTree;
  getDoc: () => Document;
  getMode: () => EditMode;
  scope: () => GeoScope | undefined;
  reviewAuthor: () => UserInfo;
  mentionableUsers: () => UserInfo[];
  addComment: (body: Fragment, mentions?: UserInfo[]) => string | null;
  focusProxy: () => void;
}

export interface CommentController {
  /** Expand the full composer at the current selection (chip click, or API). */
  openComposer(): void;
  /** Tear down the open composer, if any. */
  closeComposer(): void;
  /** Remove the floating chip, if shown. */
  hideChip(): void;
  /** Show/hide the chip as the selection/mode changes (suggest mode only). */
  updateAffordance(): void;
}

export function createCommentController(deps: CommentControllerDeps): CommentController {
  const {
    container, caretToContainer, getSelection, getCellSelection, getTree, getDoc, getMode,
    scope, reviewAuthor, mentionableUsers, addComment, focusProxy,
  } = deps;

  let bubble: HTMLDivElement | null = null;
  let chip: HTMLButtonElement | null = null;

  const closeComposer = (): void => {
    bubble?.remove();
    bubble = null;
  };
  const hideChip = (): void => {
    chip?.remove();
    chip = null;
  };

  /** Container-relative anchor (top-right of the selection's first line) for the
   *  comment chip + composer, or null if there's no usable ranged selection. */
  const anchorAt = (): { left: number; top: number; lineH: number } | null => {
    const selection = getSelection();
    if (!selection || isCollapsed(selection) || getCellSelection()) return null;
    const rects = selectionRects(getTree(), selection, scope());
    if (rects.length === 0) return null;
    const r = rects[0]!;
    const at = caretToContainer({ pageIndex: r.pageIndex, x: r.x + r.width, y: r.y, height: r.height });
    return at ? { left: at.left, top: at.top, lineH: r.height } : null;
  };

  const styleForComment = (): CharStyle => {
    const selection = getSelection();
    const block = selection ? blockById(getDoc(), selection.focus.blockId) : undefined;
    return (
      (block && selection && styleAtRuns(block.runs, selection.focus.offset)) ?? {
        fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#202124",
      }
    );
  };

  const openComposer = (): void => {
    if (getMode() === "view") return;
    const anchor = anchorAt();
    if (!anchor) return;
    hideChip();
    closeComposer();
    const style = styleForComment();

    const el = document.createElement("div");
    el.className = "cw-comment-bubble";
    const maxLeft = container.clientWidth - 312;
    el.style.left = `${Math.max(8, Math.min(anchor.left + 6, maxLeft))}px`;
    el.style.top = `${anchor.top + anchor.lineH + 8}px`;
    el.addEventListener("mousedown", (e) => e.stopPropagation());

    const who = reviewAuthor();
    const row = document.createElement("div");
    row.className = "cw-bubble-row";
    const av = document.createElement("div");
    av.className = "cw-avatar";
    av.style.background = colorForId(who.id);
    av.textContent = (who.firstName[0] ?? "?") + (who.lastName[0] ?? "");
    const ta = document.createElement("textarea");
    ta.placeholder = "Add a comment…  (@ to mention)";
    row.append(av, ta);
    const mentions = attachMentionAutocomplete(ta, mentionableUsers);

    const actions = document.createElement("div");
    actions.className = "cw-bubble-actions";
    const cancel = document.createElement("button");
    cancel.className = "cw-btn cw-btn-sm";
    cancel.textContent = "Cancel";
    const submit = document.createElement("button");
    submit.className = "cw-btn cw-btn-primary cw-btn-sm";
    submit.textContent = "Comment";
    submit.disabled = true;
    actions.append(cancel, submit);
    el.append(row, actions);
    container.appendChild(el);
    bubble = el;
    setTimeout(() => ta.focus(), 0);

    ta.addEventListener("input", () => {
      submit.disabled = ta.value.trim().length === 0;
    });
    const send = (): void => {
      const text = ta.value.trim();
      if (!text) return;
      addComment([{ text, style }], mentions.getMentions());
      mentions.destroy();
      closeComposer();
      focusProxy();
    };
    ta.addEventListener("keydown", (e) => {
      if (e.key === "Escape") { e.preventDefault(); mentions.destroy(); closeComposer(); focusProxy(); }
      else if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) { e.preventDefault(); send(); }
    });
    cancel.addEventListener("click", () => { mentions.destroy(); closeComposer(); focusProxy(); });
    submit.addEventListener("click", send);
  };

  const updateAffordance = (): void => {
    if (bubble) return; // composer open → leave it; chip is irrelevant
    const anchor = getMode() === "suggest" ? anchorAt() : null;
    if (!anchor) {
      hideChip();
      return;
    }
    if (!chip) {
      chip = document.createElement("button");
      chip.className = "cw-comment-chip";
      chip.title = "Leave a comment";
      chip.textContent = "💬";
      // stopPropagation is essential: without it the selection controller sees the
      // mousedown and collapses the selection, so openComposer finds nothing.
      chip.addEventListener("mousedown", (e) => { e.preventDefault(); e.stopPropagation(); });
      chip.addEventListener("click", (e) => { e.stopPropagation(); openComposer(); });
      container.appendChild(chip);
    }
    chip.style.left = `${anchor.left + 6}px`;
    chip.style.top = `${anchor.top - 4}px`;
  };

  return { openComposer, closeComposer, hideChip, updateAffordance };
}
