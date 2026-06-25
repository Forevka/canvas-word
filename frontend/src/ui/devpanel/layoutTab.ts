// Layout tab — the laid-out geometry tree (pages → blocks → lines → fragments).
// Wires the reusable tree view to the layout-tree builder; hovering a node paints
// its exact page rect, clicking scrolls to it.

import { buildLayoutTree, layoutSignature } from "./layoutTree";
import { createTreeView } from "./treeView";
import type { PanelCtx, PanelTab } from "./types";

export function createLayoutTab(ctx: PanelCtx): PanelTab {
  const tree = createTreeView({
    getRoots: () => buildLayoutTree(ctx.editor.getLayoutTree()),
    signature: () => layoutSignature(ctx.editor.getLayoutTree()),
    caretBlockId: () => ctx.editor.getSelection()?.focus.blockId ?? null,
    initialExpanded: ["L:page0"],
    onSelect: (node) => {
      ctx.showDetail(node.label, node.data);
      if (node.target) ctx.editor.revealInspectorTarget(node.target);
    },
    onHoverTarget: (target) => ctx.editor.setInspectorHighlight(target),
  });

  return {
    id: "layout",
    label: "Layout",
    element: tree.element,
    usesFilter: true,
    setFilter: (f) => tree.setFilter(f),
    setFollowCaret: (on) => tree.setFollowCaret(on),
    activate: () => tree.refresh(),
    refresh: () => tree.refresh(),
    highlightByBlockId: (blockId) => tree.highlightByBlockId(blockId),
    destroy: () => tree.destroy(),
  };
}
