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
import { makeFloatingDialog } from "./floatingDialog";
import { injectCssOnce } from "./styles";

/** The slice of the editor the inspector needs — the full Editor satisfies it. */
export interface DevPanelEditor {
  getDocument(): Document;
  getSelection(): DocSelection | null;
  setSelection(sel: DocSelection | null): void;
  revealBlock(blockId: string): void;
  /** Paint a devtools highlight over a node's region (null clears). */
  setInspectorHighlight(blockId: string | null): void;
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

/** One node in the rendered tree. `blockId` powers hover-highlight + click-reveal;
 *  `range` selects an inline run; `data` feeds the JSON detail pane. */
interface TreeNode {
  /** Stable key for expand-state + reverse lookup (block id where possible). */
  key: string;
  label: string;
  preview: string;
  badges: string[];
  blockId?: string | undefined;
  range?: { blockId: string; start: number; end: number } | undefined;
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

/** Build the tree node for one block, recursing into table cells. */
function blockNode(block: Block): TreeNode {
  if (block.kind === "paragraph") {
    const runs: TreeNode[] = [];
    let off = 0;
    block.runs.forEach((r, i) => {
      const start = off;
      off += r.text.length;
      runs.push({
        key: `${block.id}#run${i}`,
        label: "run",
        preview: previewText(r.text) || "∅",
        badges: runBadges(r.style),
        range: { blockId: block.id, start, end: off },
        data: r,
        children: [],
      });
    });
    return {
      key: block.id,
      label: "¶ paragraph",
      preview: previewText(textOfRuns(block.runs)) || "(empty)",
      badges: paraBadges(block),
      blockId: block.id,
      data: block,
      children: runs,
    };
  }
  if (block.kind === "image") {
    return {
      key: block.id,
      label: "🖼 image",
      preview: block.mediaId ? `media ${shortId(block.mediaId)}` : "(inline)",
      badges: imageBadges(block),
      blockId: block.id,
      data: block,
      children: [],
    };
  }
  // table → rows → cells → blocks
  const rows: TreeNode[] = block.rows.map((row, ri) => ({
    key: `${block.id}#r${ri}`,
    label: `row ${ri}`,
    preview: "",
    badges: [`${row.cells.length} cells`],
    data: row,
    children: row.cells.map((cell, ci) => ({
      key: cell.id,
      label: `cell ${ri},${ci}`,
      preview: "",
      badges: [
        ...(cell.colSpan && cell.colSpan > 1 ? [`colSpan ${cell.colSpan}`] : []),
        ...(cell.rowSpan && cell.rowSpan > 1 ? [`rowSpan ${cell.rowSpan}`] : []),
        ...(cell.shading ? ["shaded"] : []),
      ],
      data: cell,
      children: cell.blocks.map((cb) => blockNode(cb)),
    })),
  }));
  return {
    key: block.id,
    label: "▦ table",
    preview: "",
    badges: tableBadges(block),
    blockId: block.id,
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
    key,
    label,
    preview: `${entries.length}`,
    badges: [],
    data: obj,
    children: entries.map(([k, v]) => ({
      key: `${key}/${k}`,
      label: k,
      preview: previewText(typeof v === "string" ? v : JSON.stringify(v)),
      badges: [],
      data: v,
      children: [],
    })),
  };
}

/** Whole-document tree: Body, then Section bands, Footnotes, and Side-tables. */
function buildTree(doc: Document): TreeNode[] {
  const roots: TreeNode[] = [];

  roots.push({
    key: "$body",
    label: "Body",
    preview: `${doc.blocks.length} blocks`,
    badges: [],
    data: { blocks: doc.blocks.length },
    children: doc.blocks.map((b) => blockNode(b)),
  });

  const bandChildren: TreeNode[] = [];
  for (const band of BAND_CONTAINERS) {
    const blocks = doc.section[band];
    if (blocks && blocks.length > 0) {
      bandChildren.push({
        key: `$band/${band}`,
        label: band,
        preview: `${blocks.length} blocks`,
        badges: [],
        data: { band, blocks: blocks.length },
        children: blocks.map((b) => blockNode(b)),
      });
    }
  }
  if (bandChildren.length > 0) {
    roots.push({ key: "$bands", label: "Section bands", preview: `${bandChildren.length}`, badges: [], data: doc.section, children: bandChildren });
  }

  const footnotes = doc.footnotes;
  if (footnotes && Object.keys(footnotes).length > 0) {
    roots.push({
      key: "$footnotes",
      label: "Footnotes",
      preview: `${Object.keys(footnotes).length}`,
      badges: [],
      data: footnotes,
      children: Object.entries(footnotes).map(([noteId, paras]) => ({
        key: `$fn/${noteId}`,
        label: `note ${noteId}`,
        preview: previewText(paras.map((p) => textOfRuns(p.runs)).join(" ")),
        badges: [`${paras.length} ¶`],
        blockId: paras[0]?.id,
        data: paras,
        children: paras.map((p) => blockNode(p)),
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
    roots.push({ key: "$side", label: "Side tables", preview: `${sideKids.length}`, badges: [], data: {}, children: sideKids });
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
    // Reveal on canvas: a run selects its range; a block scrolls to its start.
    if (node.range) {
      editor.revealBlock(node.range.blockId);
      editor.setSelection({ anchor: { blockId: node.range.blockId, offset: node.range.start }, focus: { blockId: node.range.blockId, offset: node.range.end } });
    } else if (node.blockId) {
      editor.revealBlock(node.blockId);
    }
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
      const kindCls = node.label.includes("run") ? "cw-dev-kind run" : node.key.startsWith("$") ? "cw-dev-kind tag" : "cw-dev-kind";
      const kind = el("span", kindCls, node.label);
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
      const hoverId = node.blockId ?? node.range?.blockId ?? null;
      if (hoverId) {
        row.addEventListener("mouseenter", () => editor.setInspectorHighlight(hoverId));
        row.addEventListener("mouseleave", () => editor.setInspectorHighlight(null));
        rowByBlock.set(hoverId, row);
      }
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
