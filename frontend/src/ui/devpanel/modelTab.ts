// Model tab — the parsed-Document tree. Thin wrapper that wires the reusable tree
// view to the model-tree builder and the editor's highlight/reveal surface.

import { buildModelTree, modelSignature } from "./modelTree";
import { createTreeView } from "./treeView";
import type { PanelCtx, PanelTab } from "./types";

export function createModelTab(ctx: PanelCtx): PanelTab {
  const tree = createTreeView({
    getRoots: () => buildModelTree(ctx.editor.getDocument()),
    signature: () => modelSignature(ctx.editor.getDocument()),
    caretBlockId: () => ctx.editor.getSelection()?.focus.blockId ?? null,
    initialExpanded: ["$body"],
    onSelect: (node) => {
      ctx.showDetail(`${node.label}${node.blockId ? `  ·  ${node.blockId}` : ""}`, node.data);
      if (node.target) ctx.editor.revealInspectorTarget(node.target);
    },
    onHoverTarget: (target) => ctx.editor.setInspectorHighlight(target),
  });

  return {
    id: "model",
    label: "Model",
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
