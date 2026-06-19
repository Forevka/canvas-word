// Ribbon customization — public API types + pure ordering helpers.
//
// The editor's ribbon is built imperatively in editorApp.ts, but every built-in
// tab / group / button now registers under a stable, namespaced id (e.g. "home",
// "home.font", "home.font.bold"). An embedder passes `customizeRibbon(api)` to
// reorder/remove built-ins and add their own tabs, groups, and buttons — for
// custom macros, config popups, and informational popups.
//
// This module is DOM-free so its ordering logic is unit-testable; editorApp.ts
// owns the actual DOM wiring and implements RibbonApi against it.

import type { DocSelection } from "@cw/shared";
import type { CurrentFormat } from "./index";
import type { EditorHandle } from "./app/runtime";

/** Where to place a tab/group/button relative to an existing one (by id). When
 *  the anchor id is absent (or omitted), the item is appended at the end. */
export interface RibbonAnchor {
  before?: string;
  after?: string;
}

/** A custom ribbon button. `onClick` receives a RibbonActionContext (the editor
 *  handle + macro helpers) so it can read/edit the document or open a popup. */
export interface RibbonButtonSpec {
  /** Stable id (namespace it to avoid clashes, e.g. "myco.macros.upper"). */
  id: string;
  /** Raw innerHTML for the button face — an SVG string, emoji, or text. */
  icon?: string;
  /** Text label (used when `icon` is omitted; rendered like the B/I/U buttons). */
  label?: string;
  /** Hover tooltip. */
  tooltip?: string;
  /** Invoked on click with the editor handle + macro helpers. */
  onClick: (ctx: RibbonActionContext) => void;
  /** Optional pressed-state predicate, re-evaluated on every selection change. */
  active?: (fmt: CurrentFormat) => boolean;
  /** Optional enabled predicate; the button greys out when it returns false. */
  enabled?: (fmt: CurrentFormat) => boolean;
}

export interface RibbonTabSpec extends RibbonAnchor {
  id: string;
  label: string;
}

export interface RibbonGroupSpec extends RibbonAnchor {
  id: string;
  label: string;
}

/** The editor handle handed to a custom button's onClick, plus macro helpers. */
export interface RibbonActionContext extends EditorHandle {
  /** Current selection/caret, or null if the editor isn't focused. */
  getSelection(): DocSelection | null;
  /** Insert plain text at the caret (replacing any selection). */
  insertText(text: string): void;
  /** Emit a custom event to the embedder (`wc.on("custom", …)`). */
  emit(name: string, payload?: unknown): void;
  /** Tie a DOM node or callback to the editor's teardown so an embedder's popup
   *  is cleaned up on destroy(). Pass an element to remove, or a function to run. */
  registerCleanup(target: HTMLElement | (() => void)): void;
}

/** The mutation API passed to `customizeRibbon`. All ordering is by id; an
 *  unknown id is a no-op with a console warning (resilient to version drift). */
export interface RibbonApi {
  /** Built-in + custom tab ids in display order. */
  tabs(): string[];
  /** Group ids within a tab, in display order. */
  groups(tabId: string): string[];
  /** Item (button/control) ids within a group, in display order. */
  items(groupId: string): string[];
  addTab(spec: RibbonTabSpec): void;
  removeTab(id: string): void;
  moveTab(id: string, pos: RibbonAnchor): void;
  addGroup(tabId: string, spec: RibbonGroupSpec): void;
  removeGroup(id: string): void;
  moveGroup(id: string, pos: RibbonAnchor): void;
  addButton(groupId: string, spec: RibbonButtonSpec & RibbonAnchor): void;
  removeItem(id: string): void;
  moveItem(id: string, pos: RibbonAnchor & { toGroup?: string }): void;
}

/** Hook signature for the `customizeRibbon` constructor option. */
export type CustomizeRibbon = (api: RibbonApi) => void;

/** Index in `order` at which to insert, honoring before/after. Falls back to the
 *  end when the anchor id is missing or no anchor is given. `before` wins over
 *  `after` when both are supplied. Pure — the basis for the DOM placement. */
export function resolveInsertIndex(order: readonly string[], anchor?: RibbonAnchor): number {
  if (anchor?.before !== undefined) {
    const i = order.indexOf(anchor.before);
    if (i >= 0) return i;
  }
  if (anchor?.after !== undefined) {
    const i = order.indexOf(anchor.after);
    if (i >= 0) return i + 1;
  }
  return order.length;
}

/** The id currently at the resolved insertion index (the node to insert BEFORE in
 *  the DOM), or null to append. Skips `movingId` itself so a move computes its
 *  target against the list as if it were already removed. */
export function anchorBeforeId(
  order: readonly string[],
  anchor: RibbonAnchor | undefined,
  movingId?: string,
): string | null {
  const filtered = movingId === undefined ? [...order] : order.filter((id) => id !== movingId);
  const idx = resolveInsertIndex(filtered, anchor);
  return idx < filtered.length ? filtered[idx]! : null;
}
