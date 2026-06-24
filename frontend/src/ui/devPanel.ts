// Develop-mode Document-tree inspector — a browser-devtools-Elements-style panel
// that renders the parsed `Document` model as a navigable tree. Click a node to
// select/scroll to it on the canvas; hover a node to paint its region (devtools
// blue box); hover the canvas to reveal the matching tree node (reverse sync); and
// select a node to read its properties + raw JSON. Purely a debugging aid.
//
// Gated twice over: it only exists when the embedder passes `develop: true`, and
// even then runs nothing until the developer opens it from the Developer ribbon tab.
//
// Pattern mirrors styleManager.ts / sdtInspector.ts: injectCssOnce + makeFloatingDialog,
// an AbortController for teardown, Escape-to-close, a returned handle with
// close()/refresh(). The tree → canvas highlight + canvas → tree sync go through the
// editor's setInspectorHighlight / onInspectorHover surface (see ../index.ts).

import type { Block, DocSelection, Document, ImageBlock, Paragraph, Run, TableBlock } from "@cw/shared";
import { BAND_CONTAINERS, textOfRuns } from "@cw/shared";
import type { InspectorTarget } from "../index";
import { makeFloatingDialog } from "./floatingDialog";
import { injectCssOnce } from "./styles";

/** The slice of the editor the inspector needs — the full Editor satisfies it. */
export interface DevPanelEditor {
  getDocument(): Document;
  getSelection(): DocSelection | null;
  /** Paint a devtools highlight over a node's region (null clears). */
  setInspectorHighlight(target: InspectorTarget | null): void;
  /** Scroll a node into view (and move the caret into text targets). */
  revealInspectorTarget(target: InspectorTarget): void;
  /** Turn the canvas→tree hover signal on/off (the panel owns its lifetime). */
  setInspectorActive(active: boolean): void;
  focus(): void;
}

export interface DevPanelOptions {
  editor: DevPanelEditor;
  onClose?: () => void;
}

export interface DevPanelHandle {
  close(): void;
  /** Re-read the document and rebuild the tree, preserving expand + scroll state. */
  refresh(): void;
  /** Reverse sync: highlight (and scroll to) the tree row for a canvas block id. */
  highlightNode(blockId: string | null): void;
}

/** One node in the rendered tree. `target` is what the editor highlights/scrolls to
 *  on hover/click; `blockId` (paragraph/image/table) keys the canvas→tree reverse
 *  highlight; `data` feeds the JSON detail pane. */
interface TreeNode {
  /** Stable key for expand-state + reverse lookup (block id where possible). */
  key: string;
  label: string;
  /** Kind class for coloring: "group" (SDT/Field), "run", "tag" (structural), or "". */
  kind: "block" | "run" | "group" | "tag";
  preview: string;
  badges: string[];
  blockId?: string | undefined;
  target?: InspectorTarget | undefined;
  data: unknown;
  children: TreeNode[];
}

const CSS = `
.cw-dev-backdrop{position:fixed;inset:0;z-index:1100;}
.cw-dev-modal{position:fixed;width:min(460px,96vw);height:min(680px,92vh);min-width:300px;min-height:220px;
  max-width:96vw;max-height:96vh;display:flex;flex-direction:column;background:#1e1f22;resize:both;
  border-radius:10px;box-shadow:0 18px 56px rgba(0,0,0,.5);font:12px/1.5 'Segoe UI',Arial,sans-serif;color:#d4d6da;overflow:hidden;}
.cw-dev-head{display:flex;align-items:center;gap:8px;padding:9px 12px;border-bottom:1px solid #34363b;background:#26282c;cursor:move;}
.cw-dev-head h2{margin:0;font-size:13px;font-weight:600;flex:1 1 auto;color:#e8eaed;}
.cw-dev-x{border:none;background:transparent;font-size:18px;line-height:1;color:#9aa0a6;cursor:pointer;width:26px;height:26px;border-radius:6px;}
.cw-dev-x:hover{background:#34363b;color:#fff;}
.cw-dev-filterrow{display:flex;align-items:center;gap:6px;padding:7px 10px;border-bottom:1px solid #34363b;background:#222428;}
.cw-dev-filter{flex:1 1 auto;height:26px;padding:0 8px;border:1px solid #3a3d42;border-radius:6px;background:#1a1b1e;color:#e8eaed;font-size:12px;}
.cw-dev-filter::placeholder{color:#6b6f76;}
.cw-dev-tree{flex:1 1 auto;overflow:auto;padding:4px 0;font-family:'Cascadia Code',Consolas,monospace;}
.cw-dev-row{display:flex;align-items:baseline;gap:5px;padding:1px 10px 1px 0;cursor:pointer;white-space:nowrap;}
.cw-dev-row:hover{background:#2b2d31;}
.cw-dev-row.sel{background:#33415e;}
.cw-dev-row.cur{outline:1px solid #5b9bd5;outline-offset:-1px;}
.cw-dev-row.match{background:#3a3320;}
.cw-dev-tw{display:inline-block;width:12px;flex:0 0 auto;text-align:center;color:#9aa0a6;user-select:none;}
.cw-dev-kind{color:#9cdcfe;flex:0 0 auto;}
.cw-dev-kind.tag{color:#c586c0;}
.cw-dev-kind.run{color:#7ec699;}
.cw-dev-kind.group{color:#d7ba7d;}
.cw-dev-prev{color:#cea36a;overflow:hidden;text-overflow:ellipsis;}
.cw-dev-badge{color:#6b9bd1;font-size:10px;border:1px solid #3f5470;border-radius:7px;padding:0 5px;flex:0 0 auto;}
.cw-dev-id{color:#6b6f76;font-size:10px;flex:0 0 auto;}
.cw-dev-empty{color:#6b6f76;padding:6px 14px;font-style:italic;}
.cw-dev-detail{flex:0 0 200px;border-top:1px solid #34363b;background:#191a1d;display:flex;flex-direction:column;min-height:0;}
.cw-dev-detail-head{padding:6px 12px;border-bottom:1px solid #2b2d31;color:#9aa0a6;font-size:11px;text-transform:uppercase;letter-spacing:.05em;}
.cw-dev-json{flex:1 1 auto;overflow:auto;margin:0;padding:8px 12px;font-family:'Cascadia Code',Consolas,monospace;font-size:11px;color:#cdd1d6;white-space:pre;}`;

const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string, text?: string): HTMLElementTagNameMap[K] => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
};

const INDENT_PX = 13;
const PREVIEW_MAX = 56;

const previewText = (s: string): string => {
  const t = s.replace(/\s+/g, " ").trim();
  return t.length > PREVIEW_MAX ? `${t.slice(0, PREVIEW_MAX)}…` : t;
};

const runBadges = (style: Run["style"]): string[] => {
  const b: string[] = [];
  if (style.bold) b.push("B");
  if (style.italic) b.push("I");
  if (style.underline) b.push("U");
  if (style.strikethrough) b.push("S");
  if (style.hidden) b.push("hidden");
  if (style.verticalAlign) b.push(style.verticalAlign);
  if (style.link) b.push("link");
  if (style.footnoteRef) b.push("footnote");
  if (style.charStyleId) b.push(`cs:${style.charStyleId}`);
  if (style.sdtPath?.length) b.push(`sdt:${style.sdtPath.length}`);
  if (style.fieldId) b.push("field");
  return b;
};

const paraBadges = (p: Paragraph): string[] => {
  const b: string[] = [];
  if (p.style.namedStyle) b.push(p.style.namedStyle);
  if (p.style.list) b.push(`list L${p.style.list.level}`);
  if (p.style.outlineLevel !== undefined) b.push(`outline ${p.style.outlineLevel}`);
  if (p.sdtPath?.length) b.push(`sdt:${p.sdtPath.length}`);
  if (p.fieldId) b.push("field");
  return b;
};

const imageBadges = (im: ImageBlock): string[] => {
  const b = [`${Math.round(im.widthPx)}×${Math.round(im.heightPx)}`];
  if (im.wrap) b.push(im.wrap);
  if (im.anchor) b.push("anchored");
  if (im.sdtPath?.length) b.push(`sdt:${im.sdtPath.length}`);
  if (im.fieldId) b.push("field");
  return b;
};

const tableBadges = (t: TableBlock): string[] => {
  const cols = t.rows[0]?.cells.length ?? 0;
  const b = [`${t.rows.length}×${cols}`];
  if (t.styleId) b.push(`style:${t.styleId}`);
  if (t.sdtPath?.length) b.push(`sdt:${t.sdtPath.length}`);
  if (t.fieldId) b.push("field");
  return b;
};

const shortId = (id: string): string => (id.length > 10 ? `…${id.slice(-7)}` : id);

// ---- SDT / Field grouping -------------------------------------------------
// Content-control + field membership is stored as flat marker PATHS on blocks
// (Block.sdtPath / Block.fieldId) and runs (CharStyle.sdtPath / fieldId), not as
// tree edges. To render "SDT → paragraph → runs" and "Field → run", we reconstruct
// the grouping: a contiguous run of siblings sharing an ancestry id is wrapped in a
// group node, nesting one path level per depth. A field id is treated as the
// innermost path segment ("field:<id>"), so it nests inside any enclosing SDTs.

const blockPath = (b: Block): string[] => [...(b.sdtPath ?? []), ...(b.fieldId ? [`field:${b.fieldId}`] : [])];
const runPath = (r: Run): string[] => [...(r.style.sdtPath ?? []), ...(r.style.fieldId ? [`field:${r.style.fieldId}`] : [])];

/** A group node for one ancestry id (an SDT id, or "field:<fieldId>"). */
function groupNode(doc: Document, id: string, key: string, children: TreeNode[]): TreeNode {
  if (id.startsWith("field:")) {
    const fid = id.slice(6);
    const def = doc.fields?.[fid];
    return {
      key, kind: "group",
      label: `❏ Field ${def?.name ?? shortId(fid)}`,
      preview: previewText(def?.instruction ?? ""),
      badges: def?.kind ? [def.kind] : [],
      target: { kind: "field", fieldId: fid },
      data: def ?? { fieldId: fid },
      children,
    };
  }
  const sdt = doc.sdts?.[id];
  return {
    key, kind: "group",
    label: `❑ SDT ${sdt?.alias ?? sdt?.tag ?? sdt?.type ?? shortId(id)}`,
    preview: "",
    badges: sdt?.type ? [sdt.type] : [],
    target: { kind: "sdt", sdtId: id },
    data: sdt ?? { sdtId: id },
    children,
  };
}

/** Wrap a sequence of items into nested SDT/Field group nodes by shared ancestry.
 *  Items whose path is exhausted at this depth become leaves via `makeLeaf`. */
function groupByPath<T>(
  doc: Document,
  items: T[],
  getPath: (t: T) => string[],
  keyOf: (t: T) => string,
  makeLeaf: (t: T) => TreeNode,
  depth: number,
): TreeNode[] {
  const out: TreeNode[] = [];
  let i = 0;
  while (i < items.length) {
    const path = getPath(items[i]!);
    if (path.length <= depth) { out.push(makeLeaf(items[i]!)); i++; continue; }
    const id = path[depth]!;
    let j = i;
    while (j < items.length) {
      const p = getPath(items[j]!);
      if (p.length > depth && p[depth] === id) j++;
      else break;
    }
    const slice = items.slice(i, j);
    out.push(groupNode(doc, id, `grp/${depth}/${id}/${keyOf(slice[0]!)}`, groupByPath(doc, slice, getPath, keyOf, makeLeaf, depth + 1)));
    i = j;
  }
  return out;
}

/** Build the tree node for one block, recursing into table cells, with inline runs
 *  grouped under their content-control / field membership. */
function blockNode(doc: Document, block: Block): TreeNode {
  if (block.kind === "paragraph") {
    type RunItem = { run: Run; index: number; start: number; end: number };
    const items: RunItem[] = [];
    let off = 0;
    block.runs.forEach((r, i) => { const start = off; off += r.text.length; items.push({ run: r, index: i, start, end: off }); });
    const runLeaf = (it: RunItem): TreeNode => ({
      key: `${block.id}#run${it.index}`,
      kind: "run",
      label: "run",
      preview: previewText(it.run.text) || "∅",
      badges: runBadges(it.run.style),
      target: { kind: "run", blockId: block.id, start: it.start, end: it.end },
      data: it.run,
      children: [],
    });
    return {
      key: block.id, kind: "block",
      label: "¶ paragraph",
      preview: previewText(textOfRuns(block.runs)) || "(empty)",
      badges: paraBadges(block),
      blockId: block.id,
      target: { kind: "block", blockId: block.id },
      data: block,
      children: groupByPath(doc, items, (it) => runPath(it.run), (it) => `${block.id}#run${it.index}`, runLeaf, 0),
    };
  }
  if (block.kind === "image") {
    return {
      key: block.id, kind: "block",
      label: "🖼 image",
      preview: block.mediaId ? `media ${shortId(block.mediaId)}` : "(inline)",
      badges: imageBadges(block),
      blockId: block.id,
      target: { kind: "block", blockId: block.id },
      data: block,
      children: [],
    };
  }
  // table → rows → cells → (grouped) blocks
  const rows: TreeNode[] = block.rows.map((row, ri) => ({
    key: `${block.id}#r${ri}`, kind: "tag",
    label: `row ${ri}`,
    preview: "",
    badges: [`${row.cells.length} cells`],
    data: row,
    children: row.cells.map((cell, ci) => ({
      key: cell.id, kind: "tag",
      label: `cell ${ri},${ci}`,
      preview: previewText(cell.blocks.map((b) => (b.kind === "paragraph" ? textOfRuns(b.runs) : "")).join(" ")),
      badges: [
        ...(cell.colSpan && cell.colSpan > 1 ? [`colSpan ${cell.colSpan}`] : []),
        ...(cell.rowSpan && cell.rowSpan > 1 ? [`rowSpan ${cell.rowSpan}`] : []),
        ...(cell.shading ? ["shaded"] : []),
      ],
      target: { kind: "cell", tableId: block.id, ri, ci },
      data: cell,
      children: groupByPath(doc, cell.blocks, blockPath, (b) => b.id, (b) => blockNode(doc, b), 0),
    })),
  }));
  return {
    key: block.id, kind: "block",
    label: "▦ table",
    preview: "",
    badges: tableBadges(block),
    blockId: block.id,
    target: { kind: "block", blockId: block.id },
    data: block,
    children: rows,
  };
}

/** A non-empty record/array section as a tree node, or null when there's nothing. */
function recordNode(key: string, label: string, obj: Record<string, unknown> | undefined): TreeNode | null {
  if (!obj) return null;
  const entries = Object.entries(obj);
  if (entries.length === 0) return null;
  return {
    key, kind: "tag", label,
    preview: `${entries.length}`,
    badges: [],
    data: obj,
    children: entries.map(([k, v]) => ({
      key: `${key}/${k}`, kind: "tag",
      label: k,
      preview: previewText(typeof v === "string" ? v : JSON.stringify(v)),
      badges: [],
      data: v,
      children: [],
    })),
  };
}

/** Whole-document tree: Body, then Section bands, Footnotes, and Side-tables. Each
 *  block list is grouped by its content-control / field membership. */
function buildTree(doc: Document): TreeNode[] {
  const roots: TreeNode[] = [];
  const blockChildren = (blocks: Block[]): TreeNode[] =>
    groupByPath(doc, blocks, blockPath, (b) => b.id, (b) => blockNode(doc, b), 0);

  roots.push({
    key: "$body", kind: "tag", label: "Body",
    preview: `${doc.blocks.length} blocks`,
    badges: [],
    data: { blocks: doc.blocks.length },
    children: blockChildren(doc.blocks),
  });

  const bandChildren: TreeNode[] = [];
  for (const band of BAND_CONTAINERS) {
    const blocks = doc.section[band];
    if (blocks && blocks.length > 0) {
      bandChildren.push({
        key: `$band/${band}`, kind: "tag", label: band,
        preview: `${blocks.length} blocks`,
        badges: [],
        data: { band, blocks: blocks.length },
        children: blockChildren(blocks),
      });
    }
  }
  if (bandChildren.length > 0) {
    roots.push({ key: "$bands", kind: "tag", label: "Section bands", preview: `${bandChildren.length}`, badges: [], data: doc.section, children: bandChildren });
  }

  const footnotes = doc.footnotes;
  if (footnotes && Object.keys(footnotes).length > 0) {
    roots.push({
      key: "$footnotes", kind: "tag", label: "Footnotes",
      preview: `${Object.keys(footnotes).length}`,
      badges: [],
      data: footnotes,
      children: Object.entries(footnotes).map(([noteId, paras]) => ({
        key: `$fn/${noteId}`, kind: "tag",
        label: `note ${noteId}`,
        preview: previewText(paras.map((p) => textOfRuns(p.runs)).join(" ")),
        badges: [`${paras.length} ¶`],
        blockId: paras[0]?.id,
        data: paras,
        children: blockChildren(paras),
      })),
    });
  }

  const sideKids = [
    recordNode("$styles", "Stylesheet", doc.stylesheet as unknown as Record<string, unknown> | undefined),
    recordNode("$lists", "Lists", doc.lists),
    recordNode("$tableStyles", "Table styles", doc.tableStyles),
    recordNode("$sdts", "Content controls (SDT)", doc.sdts),
    recordNode("$fields", "Fields", doc.fields),
    recordNode("$bookmarks", "Bookmarks", doc.bookmarks),
  ].filter((n): n is TreeNode => n !== null);
  if (sideKids.length > 0) {
    roots.push({ key: "$side", kind: "tag", label: "Side tables", preview: `${sideKids.length}`, badges: [], data: {}, children: sideKids });
  }

  return roots;
}

export function showDevPanel(opts: DevPanelOptions): DevPanelHandle {
  injectCssOnce("cw-dev-styles", CSS);
  const { editor } = opts;

  // ---- DOM scaffold ---------------------------------------------------------
  const backdrop = el("div", "cw-dev-backdrop");
  const modal = el("div", "cw-dev-modal");
  modal.addEventListener("mousedown", (e) => e.stopPropagation());

  const head = el("div", "cw-dev-head");
  const h2 = el("h2", undefined, "Document Inspector");
  const xBtn = el("button", "cw-dev-x", "×");
  head.append(h2, xBtn);

  const filterRow = el("div", "cw-dev-filterrow");
  const filterInput = el("input", "cw-dev-filter");
  filterInput.type = "search";
  filterInput.placeholder = "Filter by kind, id, text, or style…";
  filterRow.append(filterInput);

  const treeHost = el("div", "cw-dev-tree");

  const detail = el("div", "cw-dev-detail");
  const detailHead = el("div", "cw-dev-detail-head", "Select a node");
  const json = el("pre", "cw-dev-json");
  detail.append(detailHead, json);

  modal.append(head, filterRow, treeHost, detail);
  backdrop.append(modal);
  document.body.append(backdrop);

  // ---- State ----------------------------------------------------------------
  const expanded = new Set<string>(["$body"]); // body open by default
  let selectedKey: string | null = null;
  let filter = "";
  let lastSig = "";
  // Map of blockId → its row element, for canvas→tree reverse highlight.
  let rowByBlock = new Map<string, HTMLElement>();
  let hoveredRow: HTMLElement | null = null;

  const matches = (node: TreeNode): boolean => {
    if (!filter) return true;
    const hay = `${node.label} ${node.preview} ${node.key} ${node.badges.join(" ")}`.toLowerCase();
    return hay.includes(filter);
  };
  // A node is kept when it (or any descendant) matches the filter.
  const keep = (node: TreeNode): boolean => matches(node) || node.children.some(keep);

  const showDetail = (node: TreeNode): void => {
    detailHead.textContent = `${node.label}${node.blockId ? `  ·  ${node.blockId}` : ""}`;
    try {
      json.textContent = JSON.stringify(node.data, null, 2);
    } catch {
      json.textContent = String(node.data);
    }
  };

  const selectNode = (node: TreeNode, rowEl: HTMLElement): void => {
    selectedKey = node.key;
    for (const r of treeHost.querySelectorAll(".cw-dev-row.sel")) r.classList.remove("sel");
    rowEl.classList.add("sel");
    showDetail(node);
    // Reveal on canvas: the editor scrolls to (and selects, for text) the node.
    if (node.target) editor.revealInspectorTarget(node.target);
  };

  // ---- Render ---------------------------------------------------------------
  const render = (): void => {
    const tree = buildTree(editor.getDocument());
    treeHost.replaceChildren();
    rowByBlock = new Map();
    const caretBlock = editor.getSelection()?.focus.blockId ?? null;

    const renderNode = (node: TreeNode, depth: number): void => {
      if (!keep(node)) return;
      const hasKids = node.children.length > 0;
      const isOpen = expanded.has(node.key) || (filter !== "" && node.children.some(keep));

      const row = el("div", "cw-dev-row");
      row.style.paddingLeft = `${6 + depth * INDENT_PX}px`;
      if (node.key === selectedKey) row.classList.add("sel");
      if (node.blockId && node.blockId === caretBlock) row.classList.add("cur");
      if (filter && matches(node)) row.classList.add("match");

      const tw = el("span", "cw-dev-tw", hasKids ? (isOpen ? "▾" : "▸") : "");
      const kind = el("span", `cw-dev-kind ${node.kind}`, node.label);
      row.append(tw, kind);
      if (node.preview) row.append(el("span", "cw-dev-prev", node.preview));
      for (const b of node.badges) row.append(el("span", "cw-dev-badge", b));
      if (node.blockId) row.append(el("span", "cw-dev-id", shortId(node.blockId)));

      if (hasKids) {
        tw.addEventListener("click", (e) => {
          e.stopPropagation();
          if (expanded.has(node.key)) expanded.delete(node.key);
          else expanded.add(node.key);
          render();
        });
      }
      row.addEventListener("click", () => selectNode(node, row));
      const target = node.target;
      if (target) {
        row.addEventListener("mouseenter", () => editor.setInspectorHighlight(target));
        row.addEventListener("mouseleave", () => editor.setInspectorHighlight(null));
      }
      if (node.blockId) rowByBlock.set(node.blockId, row);
      treeHost.append(row);

      if (hasKids && isOpen) for (const c of node.children) renderNode(c, depth + 1);
    };

    if (tree.every((n) => !keep(n))) {
      treeHost.append(el("div", "cw-dev-empty", "No nodes match the filter."));
    } else {
      for (const root of tree) renderNode(root, 0);
    }
  };

  // Skip a rebuild when nothing changed (cheap signature over the doc shape) so
  // refresh() on every keystroke stays cheap and preserves scroll position.
  const refresh = (): void => {
    const doc = editor.getDocument();
    const sig = `${doc.blocks.length}:${doc.blocks.map((b) => `${b.id}.${b.revision}`).join(",")}:${filter}:${selectedKey}:${editor.getSelection()?.focus.blockId ?? ""}:${[...expanded].sort().join(",")}`;
    if (sig === lastSig) return;
    lastSig = sig;
    render();
  };

  filterInput.addEventListener("input", () => {
    filter = filterInput.value.trim().toLowerCase();
    lastSig = ""; // force a rebuild
    render();
  });

  // ---- Reverse sync: canvas hover → highlight the tree row ------------------
  const highlightNode = (blockId: string | null): void => {
    if (hoveredRow) { hoveredRow.classList.remove("cur"); hoveredRow = null; }
    if (!blockId) return;
    const row = rowByBlock.get(blockId);
    if (row) {
      row.classList.add("cur");
      row.scrollIntoView({ block: "nearest" });
      hoveredRow = row;
    }
  };

  // ---- Lifecycle ------------------------------------------------------------
  const ac = new AbortController();
  editor.setInspectorActive(true);
  const handle: DevPanelHandle = {
    close(): void {
      editor.setInspectorActive(false);
      editor.setInspectorHighlight(null);
      backdrop.remove();
      ac.abort();
      opts.onClose?.();
    },
    refresh,
    highlightNode,
  };
  window.addEventListener("keydown", (ev: KeyboardEvent) => {
    if (ev.key === "Escape") { ev.preventDefault(); ev.stopPropagation(); handle.close(); }
  }, { capture: true, signal: ac.signal });
  makeFloatingDialog({ backdrop, modal, handle: head, signal: ac.signal, noDrag: ".cw-dev-x" });
  xBtn.addEventListener("click", () => handle.close());

  render();
  lastSig = "__init__"; // refresh() will rebuild on the first real change
  return handle;
}
