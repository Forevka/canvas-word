// SDT chooser popup — the dropdown list / combo / date picker Word shows beside a
// content control. Extracted from createEditor (index.ts) so the editor factory
// doesn't own this control's DOM construction + lifecycle. Pure presentation: the
// host injects how to read the live doc/layout and how to dispatch the resulting
// edit.

import type { Document, DocPosition } from "@cw/shared";
import { findSdtRanges, sdtAtPosition, setSdtContent, toggleSdtCheckbox } from "./commands";
import type { Command } from "./state";
import { caretRect, type CaretRect, type GeoScope } from "../layout/geometry";
import type { LayoutTree } from "../layout/layoutTree";

export interface SdtPopupDeps {
  /** Editor container the popup is appended into (positioned absolutely). */
  container: HTMLElement;
  /** Map a caret rect to container-relative coordinates (paint layer). */
  caretToContainer: (caret: CaretRect) => { left: number; top: number } | null;
  /** Live document (reassigned on every edit, so pass a getter not a value). */
  getDoc: () => Document;
  /** Live layout tree (reassigned on every relayout). */
  getTree: () => LayoutTree;
  /** Active story scope for geometry queries, if any. */
  scope: () => GeoScope | undefined;
  dispatch: (cmd: Command) => void;
  /** Return focus to the IME proxy after a choice. */
  focusProxy: () => void;
}

export interface SdtPopupController {
  /** Open the chooser for control `id` (dropdown/combo/date). */
  open(id: string): void;
  /** Close + remove the popup if open. */
  close(): void;
  /** Handle a press inside a control: toggles a checkbox immediately, or schedules
   *  the chooser to open after the caret is placed. Returns true if the press was
   *  consumed (checkbox toggle). */
  handlePress(pos: DocPosition): boolean;
}

export function createSdtPopup(deps: SdtPopupDeps): SdtPopupController {
  const { container, caretToContainer, getDoc, getTree, scope, dispatch, focusProxy } = deps;
  let popup: HTMLDivElement | null = null;

  const close = (): void => {
    popup?.remove();
    popup = null;
  };

  const open = (id: string): void => {
    close();
    const doc = getDoc();
    const props = doc.sdts?.[id];
    const range = findSdtRanges(doc, id)[0];
    if (!props || !range) return;
    const rect = caretRect(getTree(), { blockId: range.blockId, offset: range.start }, scope());
    if (!rect) return;
    const at = caretToContainer(rect);
    if (!at) return;
    const panel = document.createElement("div");
    panel.style.cssText =
      `position:absolute;left:${at.left}px;top:${at.top + rect.height + 4}px;z-index:30;` +
      "background:#fff;border:1px solid #c8c8c8;border-radius:6px;box-shadow:0 4px 12px rgba(0,0,0,.22);" +
      "font:13px Arial;min-width:160px;max-height:240px;overflow:auto;padding:4px;";
    panel.addEventListener("mousedown", (e) => {
      e.stopPropagation(); // keep the press away from the selection controller
    });
    if (props.type === "date") {
      const input = document.createElement("input");
      input.type = "date";
      input.style.cssText = "font:13px Arial;border:1px solid #c8c8c8;border-radius:4px;padding:3px 6px;";
      const pick = (): void => {
        if (!input.value) return;
        const [y, m, d] = input.value.split("-").map(Number);
        const fmt = props.dateFormat ?? "M/d/yyyy";
        const text = fmt
          .replace(/yyyy/g, String(y))
          .replace(/MM/g, String(m!).padStart(2, "0"))
          .replace(/M(?!M)/g, String(m))
          .replace(/dd/g, String(d!).padStart(2, "0"))
          .replace(/d(?!d)/g, String(d));
        dispatch(setSdtContent(id, text));
        close();
        focusProxy();
      };
      input.addEventListener("change", pick);
      panel.appendChild(input);
    } else {
      for (const item of props.listItems ?? []) {
        const row = document.createElement("div");
        row.textContent = item.display;
        row.style.cssText = "padding:4px 10px;border-radius:4px;cursor:pointer;";
        row.addEventListener("mouseenter", () => (row.style.background = "#e8eaed"));
        row.addEventListener("mouseleave", () => (row.style.background = ""));
        row.addEventListener("click", () => {
          dispatch(setSdtContent(id, item.display));
          close();
          focusProxy();
        });
        panel.appendChild(row);
      }
      if ((props.listItems ?? []).length === 0) {
        const empty = document.createElement("div");
        empty.textContent = "(no list items)";
        empty.style.cssText = "padding:4px 10px;color:#80868b;";
        panel.appendChild(empty);
      }
    }
    container.appendChild(panel);
    popup = panel;
  };

  const handlePress = (pos: DocPosition): boolean => {
    close();
    const doc = getDoc();
    const id = sdtAtPosition(doc, pos);
    const props = id ? doc.sdts?.[id] : undefined;
    if (!id || !props) return false;
    if (props.type === "checkbox") {
      dispatch(toggleSdtCheckbox(id));
      return true; // consume: the click IS the toggle
    }
    if (props.type === "dropDown" || props.type === "comboBox" || props.type === "date") {
      // Open after the controller places the caret (same frame ordering).
      requestAnimationFrame(() => open(id));
    }
    return false;
  };

  return { open, close, handlePress };
}
