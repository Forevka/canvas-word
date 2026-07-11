// Compare view — the "review differences" surface for a 2-way document compare.
// Runs diffDocuments to build a merged redline + review layer, then presents a
// VS Code-style diff panel: a scrollable list of changes on the left (each with
// accept/reject) beside a live redline preview on the right (insertions
// underlined, deletions struck, change bars in the margin — the existing review
// paint channel). Resolving a change re-diffs the local state and re-renders;
// nothing touches the live editor until the user applies the result.
//
// It reuses the isomorphic resolvers (acceptAll/rejectAll/bakeReview) and the
// frontend render stack (layout engine + paint layer + decorate) exactly as the
// live editor's review pane does — no second Editor instance.

import {
  acceptAllSuggestions,
  applyOp,
  applyReviewOp,
  bakeReview,
  blockById,
  blockIndexOf,
  diffDocuments,
  gcStructuralReviewLayer,
  rebaseReview,
  rejectAllSuggestions,
  sliceRuns,
  textOfRuns,
  type DocPosition,
  type Document,
  type ReviewLayer,
  type Suggestion,
} from "@cw/shared";
import { createDialogShell } from "./dialogShell";
import { createLayoutEngine } from "../layout/engine";
import { createPaintLayer, type PaintScheduler } from "../paint/renderer";
import { decorate } from "../review/decorate";
import { caretRect } from "../layout/geometry";
import type { LayoutTree } from "../layout/layoutTree";
import type { ResolvedTheme } from "../config";
import type { CustomFontRegistry } from "../fonts/customRegistry";

export interface CompareViewOptions {
  /** The document currently open in the editor — the ORIGINAL / baseline. */
  original: Document;
  /** The picked document to compare against — the REVISED version. */
  revised: Document;
  /** Human labels for the two sides (file names). */
  originalName?: string;
  revisedName?: string;
  fontRegistry?: CustomFontRegistry;
  theme?: ResolvedTheme;
  /** Apply the result back to the host editor. `tracked` = load the merged doc +
   *  remaining suggestions in suggestion mode; `accepted` = load the plain doc
   *  with every remaining change taken. */
  onApply(doc: Document, review: ReviewLayer, mode: "tracked" | "accepted"): void;
  onClose?(): void;
}

export interface CompareViewHandle {
  close(): void;
}

/** One row in the change list — a whole group (a paragraph add/remove bundles its
 *  structural + text records under one groupId) resolved as a unit. */
interface ChangeGroup {
  key: string;
  ids: string[];
  kind: "insert" | "delete" | "replace" | "format" | "addBlock" | "removeBlock" | "paraFormat";
  label: string;
  preview: string;
  anchor: DocPosition | null;
  order: number;
}

const KIND_META: Record<ChangeGroup["kind"], { tag: string; cls: string }> = {
  insert: { tag: "Inserted", cls: "ins" },
  delete: { tag: "Deleted", cls: "del" },
  replace: { tag: "Replaced", cls: "rep" },
  format: { tag: "Formatting", cls: "fmt" },
  addBlock: { tag: "Added ¶", cls: "ins" },
  removeBlock: { tag: "Removed ¶", cls: "del" },
  paraFormat: { tag: "¶ format", cls: "fmt" },
};

const CSS = `
.cw-cmp-backdrop{position:fixed;inset:0;z-index:2147483000;pointer-events:none;}
.cw-cmp-modal{position:fixed;pointer-events:auto;width:min(96vw,1200px);height:86vh;display:flex;flex-direction:column;
  background:#fff;color:#201f1e;border:1px solid #c8c6c4;border-radius:8px;box-shadow:0 12px 40px rgba(0,0,0,.3);
  font:13px 'Segoe UI',Roboto,sans-serif;}
.cw-cmp-head{display:flex;align-items:center;gap:10px;padding:10px 12px;border-bottom:1px solid #edebe9;cursor:move;}
.cw-cmp-head h2{margin:0;font-size:15px;font-weight:600;flex:1;}
.cw-cmp-count{font-size:12px;color:#605e5c;background:#f3f2f1;border-radius:10px;padding:2px 10px;}
.cw-cmp-x{border:0;background:transparent;font-size:20px;line-height:1;cursor:pointer;color:#605e5c;padding:2px 6px;border-radius:4px;}
.cw-cmp-x:hover{background:#f3f2f1;color:#201f1e;}
.cw-cmp-body{display:flex;flex:1 1 auto;min-height:0;}
.cw-cmp-list{flex:0 0 320px;overflow:auto;border-right:1px solid #edebe9;background:#faf9f8;}
.cw-cmp-row{display:flex;flex-direction:column;gap:4px;padding:8px 10px;border-bottom:1px solid #efedeb;cursor:pointer;}
.cw-cmp-row:hover{background:#f3f2f1;}
.cw-cmp-row.active{background:#eef0ff;}
.cw-cmp-rowtop{display:flex;align-items:center;gap:6px;}
.cw-cmp-tag{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.03em;padding:1px 6px;border-radius:4px;color:#fff;white-space:nowrap;}
.cw-cmp-tag.ins{background:#107c41;}
.cw-cmp-tag.del{background:#a4262c;}
.cw-cmp-tag.rep{background:#8661c5;}
.cw-cmp-tag.fmt{background:#797673;}
.cw-cmp-prev{flex:1;color:#323130;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.cw-cmp-prev .old{color:#a4262c;text-decoration:line-through;}
.cw-cmp-prev .new{color:#107c41;}
.cw-cmp-acts{display:flex;gap:6px;}
.cw-cmp-mini{border:1px solid #c8c6c4;background:#fff;border-radius:4px;padding:2px 10px;font-size:12px;cursor:pointer;line-height:1.4;}
.cw-cmp-mini.acc:hover{background:#dff6e6;border-color:#107c41;color:#0b5a2f;}
.cw-cmp-mini.rej:hover{background:#fde7e9;border-color:#a4262c;color:#7a1c22;}
.cw-cmp-empty{padding:28px 16px;text-align:center;color:#605e5c;}
.cw-cmp-preview{flex:1 1 auto;min-width:0;overflow:auto;background:#e6e6e6;padding:16px;}
.cw-cmp-foot{display:flex;align-items:center;gap:8px;padding:10px 14px;border-top:1px solid #edebe9;}
.cw-cmp-foot .spacer{flex:1;}
.cw-cmp-btn{border:1px solid #c8c6c4;background:#fff;border-radius:4px;padding:7px 16px;font-size:13px;cursor:pointer;}
.cw-cmp-btn:hover{background:#f3f2f1;}
.cw-cmp-btn.primary{background:#6b3fa0;border-color:#6b3fa0;color:#fff;}
.cw-cmp-btn.primary:hover{background:#5a3488;}
`;

const truncate = (s: string, n = 42): string => (s.length > n ? s.slice(0, n - 1) + "…" : s);

/** Text covered by a suggestion's anchor in the given doc (for the list preview). */
function anchorText(doc: Document, s: Suggestion): string {
  const { start, end } = s.anchor;
  const block = blockById(doc, start.blockId);
  if (!block) return "";
  if (start.blockId === end.blockId && end.offset > start.offset) {
    return textOfRuns(sliceRuns(block.runs, start.offset, end.offset));
  }
  return "";
}

export function showCompareView(opts: CompareViewOptions): CompareViewHandle {
  const shell = createDialogShell({
    prefix: "cw-cmp",
    cssId: "cw-cmp-css",
    css: CSS,
    title: `Compare — ${opts.originalName ?? "Original"} ↔ ${opts.revisedName ?? "Revised"}`,
    extraNoDrag: ".cw-cmp-body, .cw-cmp-foot",
    onClose: () => {
      painter?.destroy();
      opts.onClose?.();
    },
  });

  const diff = diffDocuments(opts.original, opts.revised, { docId: "compare", startTime: Date.now() });
  let curDoc: Document = diff.mergedDoc;
  let curReview: ReviewLayer = diff.review;

  // ---- header change counter ----
  const count = document.createElement("span");
  count.className = "cw-cmp-count";
  shell.head.insertBefore(count, shell.closeBtn);

  // ---- body: change list | live preview ----
  shell.body.className = "cw-cmp-body";
  const list = document.createElement("div");
  list.className = "cw-cmp-list";
  const preview = document.createElement("div");
  preview.className = "cw-cmp-preview";
  shell.body.append(list, preview);

  const engine = createLayoutEngine(opts.fontRegistry);
  const painter: PaintScheduler = createPaintLayer(preview, {
    chrome: false,
    ...(opts.theme ? { theme: opts.theme } : {}),
    ...(opts.fontRegistry ? { fontRegistry: opts.fontRegistry } : {}),
  });

  let tree: LayoutTree = engine.layout(curDoc);
  let activeKey: string | null = null;

  // ---- resolution (accept/reject a group against the LOCAL merged state) ----
  const resolveGroup = (group: ChangeGroup, accept: boolean): void => {
    const idset = new Set(group.ids);
    const subset: ReviewLayer = { ...curReview, suggestions: curReview.suggestions.filter((s) => idset.has(s.id)) };
    const res = accept ? acceptAllSuggestions(curDoc, subset) : rejectAllSuggestions(curDoc, subset);
    // Apply the resulting core ops to the local merged doc, rebasing EVERY other
    // record's anchors through each op (mirrors the runtime commit pipeline), then
    // drop the resolved records.
    let d = curDoc;
    let rv = curReview;
    for (const op of res.ops) {
      const a = applyOp(d, op);
      d = a.doc;
      rv = gcStructuralReviewLayer(rebaseReview(rv, a.mapPosition), d);
    }
    for (const rop of res.reviewOps) rv = applyReviewOp(rv, rop).layer;
    curDoc = d;
    curReview = rv;
    render();
  };

  const resolveAll = (accept: boolean): void => {
    curDoc = bakeReview(curDoc, curReview, accept ? "accept" : "reject");
    curReview = { ...curReview, suggestions: [] };
    render();
  };

  // ---- change-list grouping ----
  const buildGroups = (): ChangeGroup[] => {
    const byKey = new Map<string, Suggestion[]>();
    for (const s of curReview.suggestions) {
      const key = s.groupId ?? s.id;
      (byKey.get(key) ?? byKey.set(key, []).get(key)!).push(s);
    }
    const groups: ChangeGroup[] = [];
    for (const [key, members] of byKey) {
      const structural = members.find((m) => m.kind === "structural" && m.structural);
      const anchorSug = members.find((m) => m.anchor.start.offset !== m.anchor.end.offset) ?? members[0]!;
      const anchor = anchorSug.anchor.start;

      let kind: ChangeGroup["kind"];
      let preview = "";
      if (structural?.structural) {
        const op = structural.structural.op.type;
        const blk = blockById(curDoc, structural.structural.blockId);
        const blkText = blk ? textOfRuns(blk.runs) : "";
        if (op === "insertBlock") {
          kind = "addBlock";
          preview = truncate(blkText);
        } else if (op === "removeBlock") {
          kind = "removeBlock";
          preview = truncate(blkText);
        } else {
          kind = "paraFormat";
          preview = truncate(blkText);
        }
      } else {
        const hasIns = members.some((m) => m.kind === "insert");
        const hasDel = members.some((m) => m.kind === "delete");
        if (hasIns && hasDel) kind = "replace";
        else if (hasIns) kind = "insert";
        else if (hasDel) kind = "delete";
        else kind = "format";
        const ins = members.find((m) => m.kind === "insert");
        const del = members.find((m) => m.kind === "delete");
        preview = truncate([del ? anchorText(curDoc, del) : "", ins ? anchorText(curDoc, ins) : ""].filter(Boolean).join(" → ") || anchorText(curDoc, anchorSug));
      }

      const bi = anchor ? blockIndexOf(curDoc, anchor.blockId) : -1;
      groups.push({ key, ids: members.map((m) => m.id), kind, label: KIND_META[kind].tag, preview, anchor, order: (bi < 0 ? 1e6 : bi) * 1e7 + (anchor?.offset ?? 0) });
    }
    groups.sort((a, b) => a.order - b.order);
    return groups;
  };

  const scrollTo = (anchor: DocPosition | null): void => {
    if (!anchor) return;
    const cr = caretRect(tree, anchor);
    if (!cr) return;
    const at = painter.caretToContainer(cr);
    if (at) preview.scrollTop = Math.max(0, at.top - 90);
  };

  const renderList = (groups: ChangeGroup[]): void => {
    list.textContent = "";
    if (groups.length === 0) {
      const empty = document.createElement("div");
      empty.className = "cw-cmp-empty";
      empty.textContent = diff.review.suggestions.length === 0 ? "The documents are identical." : "All changes resolved. Apply to finish.";
      list.appendChild(empty);
      return;
    }
    for (const g of groups) {
      const row = document.createElement("div");
      row.className = "cw-cmp-row" + (g.key === activeKey ? " active" : "");
      const top = document.createElement("div");
      top.className = "cw-cmp-rowtop";
      const tag = document.createElement("span");
      tag.className = `cw-cmp-tag ${KIND_META[g.kind].cls}`;
      tag.textContent = g.label;
      const prev = document.createElement("span");
      prev.className = "cw-cmp-prev";
      prev.textContent = g.preview;
      top.append(tag, prev);
      const acts = document.createElement("div");
      acts.className = "cw-cmp-acts";
      const acc = document.createElement("button");
      acc.className = "cw-cmp-mini acc";
      acc.textContent = "✓ Accept";
      acc.addEventListener("click", (e) => {
        e.stopPropagation();
        resolveGroup(g, true);
      });
      const rej = document.createElement("button");
      rej.className = "cw-cmp-mini rej";
      rej.textContent = "✗ Reject";
      rej.addEventListener("click", (e) => {
        e.stopPropagation();
        resolveGroup(g, false);
      });
      acts.append(acc, rej);
      row.append(top, acts);
      row.addEventListener("click", () => {
        activeKey = g.key;
        scrollTo(g.anchor);
        [...list.children].forEach((c) => c.classList.remove("active"));
        row.classList.add("active");
      });
      list.appendChild(row);
    }
  };

  const render = (): void => {
    engine.reset();
    tree = engine.layout(curDoc);
    painter.setTree(tree);
    painter.setReviewDecorations(decorate(curReview, tree));
    const groups = buildGroups();
    const n = groups.length;
    count.textContent = n === 0 ? "No changes" : n === 1 ? "1 change" : `${n} changes`;
    renderList(groups);
  };

  // ---- footer actions ----
  const mkBtn = (label: string, primary: boolean, onClick: () => void): HTMLButtonElement => {
    const b = document.createElement("button");
    b.className = "cw-cmp-btn" + (primary ? " primary" : "");
    b.textContent = label;
    b.addEventListener("click", onClick);
    return b;
  };
  shell.foot.className = "cw-cmp-foot";
  shell.foot.append(
    mkBtn("Reject all", false, () => resolveAll(false)),
    mkBtn("Accept all", false, () => resolveAll(true)),
  );
  const spacer = document.createElement("div");
  spacer.className = "spacer";
  shell.foot.append(
    spacer,
    mkBtn("Cancel", false, () => shell.close()),
    mkBtn("Apply as tracked changes", false, () => {
      opts.onApply(curDoc, curReview, "tracked");
      shell.close();
    }),
    mkBtn("Apply result", true, () => {
      opts.onApply(bakeReview(curDoc, curReview, "accept"), { ...curReview, suggestions: [] }, "accepted");
      shell.close();
    }),
  );

  render();
  return { close: () => shell.close() };
}
