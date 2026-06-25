// Problems tab — runs the model validator and lists diagnostics grouped by severity,
// each linking to the offending block. Reuses the tree view (filter + reveal + JSON).

import { modelSignature } from "./modelTree";
import { createTreeView } from "./treeView";
import type { PanelCtx, PanelTab, TreeNode } from "./types";
import { countBySeverity, validateDocument, type Diagnostic, type Severity } from "./validate";

const ICON: Record<Severity, string> = { error: "⛔", warning: "⚠", info: "ℹ" };
const GROUPS: { sev: Severity; label: string }[] = [
  { sev: "error", label: "Errors" },
  { sev: "warning", label: "Warnings" },
  { sev: "info", label: "Info" },
];

function diagnosticsToTree(diags: Diagnostic[]): TreeNode[] {
  if (diags.length === 0) {
    return [{ key: "$ok", kind: "tag", label: "✓ No problems", preview: "the document model is clean", badges: [], data: { ok: true }, children: [] }];
  }
  const roots: TreeNode[] = [];
  for (const g of GROUPS) {
    const items = diags.filter((d) => d.severity === g.sev);
    if (items.length === 0) continue;
    roots.push({
      key: `$sev/${g.sev}`, kind: "tag",
      label: `${ICON[g.sev]} ${g.label}`,
      preview: `${items.length}`,
      badges: [],
      data: { count: items.length },
      children: items.map((d, i) => ({
        key: `prob/${g.sev}/${i}`, kind: "problem",
        label: d.code,
        preview: d.message,
        badges: [],
        blockId: d.blockId,
        target: d.blockId ? { kind: "block", blockId: d.blockId } : undefined,
        data: d,
        children: [],
      })),
    });
  }
  return roots;
}

export function createProblemsTab(ctx: PanelCtx): PanelTab {
  const tree = createTreeView({
    getRoots: () => diagnosticsToTree(validateDocument(ctx.editor.getDocument())),
    signature: () => modelSignature(ctx.editor.getDocument()),
    caretBlockId: () => ctx.editor.getSelection()?.focus.blockId ?? null,
    initialExpanded: ["$sev/error", "$sev/warning", "$sev/info"],
    onSelect: (node) => {
      ctx.showDetail(node.label, node.data);
      if (node.target) ctx.editor.revealInspectorTarget(node.target);
    },
    onHoverTarget: (target) => ctx.editor.setInspectorHighlight(target),
  });

  return {
    id: "problems",
    label: "Problems",
    element: tree.element,
    usesFilter: true,
    setFilter: (f) => tree.setFilter(f),
    activate: () => tree.refresh(),
    refresh: () => tree.refresh(),
    highlightByBlockId: (blockId) => tree.highlightByBlockId(blockId),
    count: () => {
      const c = countBySeverity(validateDocument(ctx.editor.getDocument()));
      return c.error + c.warning || null;
    },
    destroy: () => tree.destroy(),
  };
}
