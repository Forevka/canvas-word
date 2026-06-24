// Shared types + tiny DOM helpers for the develop-mode inspector panel. The panel
// is a tabbed shell (Model · Layout · Problems · History) over a small editor slice;
// each tab is a `PanelTab` mounted into the shell. See devPanel.ts for the shell.

import type { Change, DocSelection, Document } from "@cw/shared";
import type { InspectorTarget } from "../../index";
import type { LayoutTree } from "../../layout/layoutTree";

/** The slice of the editor the inspector needs — the full Editor satisfies it. */
export interface DevPanelEditor {
  getDocument(): Document;
  getSelection(): DocSelection | null;
  /** The laid-out geometry tree (Layout tab + overlays). */
  getLayoutTree(): LayoutTree;
  /** Paint a devtools highlight over a node's region (null clears). */
  setInspectorHighlight(target: InspectorTarget | null): void;
  /** Scroll a node into view (and move the caret into text targets). */
  revealInspectorTarget(target: InspectorTarget): void;
  /** Turn the canvas→tree hover signal on/off (the panel owns its lifetime). */
  setInspectorActive(active: boolean): void;
  /** Turn the hit-test probe signal on/off (Probe tab owns its lifetime). */
  setInspectorProbe(active: boolean): void;
  /** Toggle a layout-debug overlay drawn on the canvas. */
  setDebugOverlay(kind: string, on: boolean): void;
  /** Recorded edit history (History tab). */
  getChangeLog(): Change[];
  getChangeHead(): number;
  undo(): void;
  redo(): void;
  focus(): void;
}

/** One node in a rendered tree. `target` is what the editor highlights/scrolls to on
 *  hover/click; `blockId` (paragraph/image/table) keys the canvas→tree reverse
 *  highlight; `data` feeds the shared detail/JSON pane. */
export interface TreeNode {
  /** Stable key for expand-state + reverse lookup (block id where possible). */
  key: string;
  label: string;
  /** Kind class for coloring. */
  kind: "block" | "run" | "group" | "tag" | "layout" | "problem";
  preview: string;
  badges: string[];
  blockId?: string | undefined;
  target?: InspectorTarget | undefined;
  data: unknown;
  children: TreeNode[];
}

/** Context handed to each tab factory: the editor slice + the shared detail pane. */
export interface PanelCtx {
  editor: DevPanelEditor;
  /** Show a node/object in the shared detail/JSON pane. */
  showDetail(title: string, data: unknown): void;
}

/** A tab in the inspector shell. The shell shows/hides `element`, swaps the toolbar
 *  extras, and drives `activate`/`refresh`/`highlightByBlockId` on the active tab. */
export interface PanelTab {
  id: string;
  label: string;
  /** The tab's content element (added once; shown/hidden on switch). */
  readonly element: HTMLElement;
  /** Extra toolbar controls shown only while this tab is active. */
  readonly toolbar?: HTMLElement | undefined;
  /** Whether the shared filter input applies (and shows) for this tab. */
  readonly usesFilter?: boolean | undefined;
  setFilter?(f: string): void;
  /** Follow-caret toggle (tree tabs scroll to the caret's node). */
  setFollowCaret?(on: boolean): void;
  /** Called when the tab becomes active (and on first show). */
  activate(): void;
  /** Called when the tab is switched away from (turn off any live signals). */
  deactivate?(): void;
  /** Live refresh on doc/selection change — only called while active. */
  refresh(): void;
  /** Reverse sync from a canvas-hover blockId (tree tabs). */
  highlightByBlockId?(blockId: string | null): void;
  /** Receive a hit-test probe payload (Probe tab). Typed as unknown to keep the
   *  shell decoupled; the Probe tab casts it. */
  onProbe?(probe: unknown): void;
  /** Optional count badge shown next to the tab label (e.g. problem count). Null = none. */
  count?(): number | null;
  destroy?(): void;
}

export const INDENT_PX = 13;
const PREVIEW_MAX = 56;

export const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string, text?: string): HTMLElementTagNameMap[K] => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
};

export const previewText = (s: string): string => {
  const t = s.replace(/\s+/g, " ").trim();
  return t.length > PREVIEW_MAX ? `${t.slice(0, PREVIEW_MAX)}…` : t;
};

export const shortId = (id: string): string => (id.length > 10 ? `…${id.slice(-7)}` : id);
