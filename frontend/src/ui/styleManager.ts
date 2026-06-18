// Style Manager — the unified, non-blocking, draggable dialog for creating,
// editing and deleting every kind of document style (Word's "Manage Styles").
// Three columns: a left rail listing styles grouped by kind, a center editor
// pane (per-kind, swapped via a small registry), and a right live preview painted
// with the real layout engine through a child document. Edits resolve against a
// DRAFT stylesheet so the preview updates before the change is committed; on Apply
// the spec is turned into an editor command (upsertNamedStyle / …).
//
// Pattern mirrors fieldConstructor.ts: injectCssOnce + makeFloatingDialog, an
// AbortController for teardown, Escape-to-close. List + Table kinds are scaffolded
// here and populated by their own phases.

import type { Block, CharStyle, Document, ListDefinition, NamedStyle, ParaStyle, Paragraph, Run, Stylesheet, TableStyle } from "@cw/shared";
import { bulletListDefinition, defaultStylesheet, descendantsOf, mergeDuplicateNamedStyles, resolveCharStyle, resolveStyle, styleById, styleType } from "@cw/shared";
import type { ChildDocument, ChildContent } from "../child/childDocument";
import type { Command } from "../editor/state";
import { deleteListDefinition, deleteNamedStyle, deleteTableStyle, mergeDuplicateStyles, setDefaultStyle, upsertListDefinition, upsertNamedStyle, upsertTableStyle, type NamedStyleSpec } from "../editor/commands";
import { showContextMenu, type MenuEntry } from "./contextMenu";
import { makeFloatingDialog } from "./floatingDialog";
import { injectCssOnce } from "./styles";
import { mountTextStyleEditor, type StyleEditorController } from "./styleEditorText";
import { mountListStyleEditor, type ListStyleController } from "./styleEditorList";
import { mountTableStyleEditor, type TableStyleController } from "./styleEditorTable";

export type StyleKind = "paragraph" | "character" | "list" | "table";

/** The slice of the editor the Style Manager needs — the full Editor satisfies it,
 *  and the editor core can pass a lightweight adapter from its context menu. */
export interface StyleManagerEditor {
  getDocument(): Document;
  dispatch(cmd: Command): void;
  createChild(): ChildDocument;
  focus(): void;
}

export interface StyleManagerOptions {
  editor: StyleManagerEditor;
  /** Select this style when the dialog opens. */
  initialSelection?: { kind: StyleKind; id: string };
  onClose?: () => void;
}

export interface StyleManagerHandle {
  close(): void;
  /** Re-read the document and rebuild the list (e.g. after an external change). */
  refresh(): void;
}

const KIND_LABELS: Record<StyleKind, string> = {
  paragraph: "Paragraph styles",
  character: "Character styles",
  list: "List styles",
  table: "Table styles",
};

const FALLBACK_CHAR: CharStyle = {
  fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false,
  underline: false, strikethrough: false, color: "#202124",
};
const FALLBACK_PARA: ParaStyle = {
  align: "left", lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0,
};
/** A thin grey edge for a new table style's whole-table grid. */
const B = { color: "#bfbfbf", widthPx: 1 } as const;

const CSS = `
.cw-sm-backdrop{position:fixed;inset:0;z-index:1100;}
.cw-sm-modal{position:fixed;width:min(980px,96vw);height:min(640px,92vh);display:flex;flex-direction:column;background:#fff;border-radius:10px;
  box-shadow:0 18px 56px rgba(0,0,0,.34);font:13px/1.5 Arial,sans-serif;color:#202124;overflow:hidden;}
.cw-sm-head{display:flex;align-items:center;gap:10px;padding:13px 16px;border-bottom:1px solid #e6e8eb;background:#f7f8fa;cursor:move;}
.cw-sm-head h2{margin:0;font-size:15px;font-weight:600;flex:1 1 auto;}
.cw-sm-x{border:none;background:transparent;font-size:20px;line-height:1;color:#5f6368;cursor:pointer;width:28px;height:28px;border-radius:6px;}
.cw-sm-x:hover{background:#e8eaed;}
.cw-sm-body{display:flex;min-height:0;flex:1 1 auto;}
.cw-sm-list{flex:0 0 230px;border-right:1px solid #e6e8eb;overflow:auto;padding:8px;background:#fbfbfc;}
.cw-sm-sec{margin-bottom:8px;}
.cw-sm-sec>h4{margin:6px 4px;font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:#9aa0a6;}
.cw-sm-card{display:flex;flex-direction:column;gap:3px;width:100%;text-align:left;border:1px solid transparent;border-radius:7px;background:transparent;cursor:pointer;padding:6px 8px;margin-bottom:2px;}
.cw-sm-card:hover{background:#f1f3f4;}
.cw-sm-card.sel{background:#e8f0fe;border-color:#1a73e8;}
.cw-sm-card .nm{font-size:12px;color:#202124;display:flex;align-items:center;gap:6px;}
.cw-sm-card .nm .def{font-size:9px;color:#1a73e8;border:1px solid #1a73e8;border-radius:8px;padding:0 5px;}
.cw-sm-card .sw{min-height:20px;overflow:hidden;}
.cw-sm-empty{font-size:11px;color:#9aa0a6;padding:6px 8px;}
.cw-sm-editor{flex:1 1 auto;overflow:auto;padding:16px;}
.cw-sm-preview{flex:0 0 270px;border-left:1px solid #e6e8eb;background:#fbfbfc;padding:16px;display:flex;flex-direction:column;gap:8px;overflow:auto;}
.cw-sm-preview .ttl{font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:#80868b;}
.cw-sm-prevhost{border:1px solid #dadce0;border-radius:8px;background:#fff;padding:8px;min-height:60px;}
.cw-sm-foot{display:flex;align-items:center;gap:8px;padding:11px 16px;border-top:1px solid #e6e8eb;}
.cw-sm-foot .spacer{flex:1 1 auto;}
.cw-sm-btn{height:30px;padding:0 14px;border:1px solid #d0d4d9;border-radius:6px;background:#fff;cursor:pointer;font-size:13px;color:#3c4043;}
.cw-sm-btn:hover{background:#f1f3f4;}
.cw-sm-btn:disabled{opacity:.5;cursor:default;}
.cw-sm-btn.primary{border-color:#1a73e8;background:#1a73e8;color:#fff;}
.cw-sm-btn.primary:hover{background:#1864cc;}
.cw-sm-btn.danger:hover{background:#fce8e6;border-color:#d93025;color:#d93025;}`;

const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string, text?: string): HTMLElementTagNameMap[K] => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
};

export function showStyleManager(opts: StyleManagerOptions): StyleManagerHandle {
  injectCssOnce("cw-sm-styles", CSS);
  const { editor } = opts;
  const child = editor.createChild();

  const sheet = (): Stylesheet => editor.getDocument().stylesheet ?? defaultStylesheet();

  // ---- DOM scaffold ---------------------------------------------------------
  const backdrop = el("div", "cw-sm-backdrop");
  const modal = el("div", "cw-sm-modal");
  modal.addEventListener("mousedown", (e) => e.stopPropagation());

  const head = el("div", "cw-sm-head");
  const h2 = el("h2", undefined, "Manage Styles");
  const xBtn = el("button", "cw-sm-x", "×");
  head.append(h2, xBtn);

  const body = el("div", "cw-sm-body");
  const listHost = el("div", "cw-sm-list");
  const editorHost = el("div", "cw-sm-editor");
  const previewCol = el("div", "cw-sm-preview");
  const prevHost = el("div", "cw-sm-prevhost");
  previewCol.append(el("div", "ttl", "Preview"), prevHost);
  body.append(listHost, editorHost, previewCol);

  const foot = el("div", "cw-sm-foot");
  const newBtn = el("button", "cw-sm-btn", "New ▾");
  const delBtn = el("button", "cw-sm-btn danger", "Delete");
  const mergeBtn = el("button", "cw-sm-btn", "Merge duplicates");
  mergeBtn.title = "Collapse styles that are identical except for their name";
  const applyBtn = el("button", "cw-sm-btn primary", "Apply");
  const closeBtn = el("button", "cw-sm-btn", "Close");
  foot.append(newBtn, delBtn, mergeBtn, el("div", "spacer"), applyBtn, closeBtn);

  modal.append(head, body, foot);
  backdrop.append(modal);
  document.body.append(backdrop);

  // ---- State ----------------------------------------------------------------
  let selKind: StyleKind = "paragraph";
  let selId: string | null = null; // null = an unsaved new style
  type Held =
    | { kind: "text"; ctl: StyleEditorController }
    | { kind: "list"; ctl: ListStyleController }
    | { kind: "table"; ctl: TableStyleController };
  let held: Held | null = null;

  // ---- Preview --------------------------------------------------------------
  const sampleBlocks = (spec: NamedStyleSpec): Block[] => {
    const s = sheet();
    const draftId = spec.id ?? "__cw_draft__";
    const draft: NamedStyle = { id: draftId, name: spec.name || "Sample", type: spec.type, char: spec.char ?? {}, para: spec.para ?? {} };
    if (spec.basedOn) draft.basedOn = spec.basedOn;
    const draftSheet: Stylesheet = { ...s, styles: [...s.styles.filter((x) => x.id !== draftId), draft] };
    const def = resolveStyle(draftSheet, draftSheet.defaultStyleId);
    const para: ParaStyle = { ...FALLBACK_PARA, ...def.para };
    const mkPara = (runs: Run[]): Paragraph => ({ kind: "paragraph", id: "cw-sm-prev", revision: 0, runs, style: para });
    if (spec.type === "character") {
      const baseChar: CharStyle = { ...FALLBACK_CHAR, ...def.char };
      const styled: CharStyle = { ...baseChar, ...resolveCharStyle(draftSheet, draftId) };
      return [mkPara([
        { text: "Normal text, ", style: baseChar },
        { text: "styled sample", style: styled },
        { text: ", normal text.", style: baseChar },
      ])];
    }
    const r = resolveStyle(draftSheet, draftId);
    const char: CharStyle = { ...FALLBACK_CHAR, ...r.char };
    const pPara: ParaStyle = { ...FALLBACK_PARA, ...r.para };
    return [{ kind: "paragraph", id: "cw-sm-prev", revision: 0, style: pPara, runs: [{ text: "The quick brown fox jumps over the lazy dog. 0123456789", style: char }] }];
  };

  const refreshPreview = (): void => {
    if (!held) { prevHost.replaceChildren(); return; }
    try {
      const content: ChildContent =
        held.kind === "list"
          ? { kind: "listSample", listDef: held.ctl.read() }
          : held.kind === "table"
            ? { kind: "tableSample", tableStyle: held.ctl.read() }
            : { kind: "blocks", blocks: sampleBlocks(held.ctl.read()) };
      child.render(prevHost, content, { padding: 8 });
    } catch {
      /* preview is best-effort */
    }
  };

  // ---- Editor pane ----------------------------------------------------------
  const basedOnOptions = (kind: "paragraph" | "character", currentId: string | null): { id: string; name: string }[] => {
    const s = sheet();
    const excl = currentId ? descendantsOf(s, currentId) : new Set<string>();
    return s.styles
      .filter((x) => styleType(x) === kind && x.id !== currentId && !excl.has(x.id))
      .map((x) => ({ id: x.id, name: x.name }));
  };

  const lists = (): Record<string, ListDefinition> => editor.getDocument().lists ?? {};
  const tableStyles = (): Record<string, TableStyle> => editor.getDocument().tableStyles ?? {};

  const defaultTableStyle = (id: string): TableStyle => ({
    id,
    name: "New Table Style",
    rowBandSize: 1,
    conds: {
      wholeTable: { borders: { top: B, right: B, bottom: B, left: B }, margin: { top: 1, right: 6, bottom: 1, left: 6 } },
      firstRow: { shading: "#4472c4", char: { bold: true, color: "#ffffff" } },
      band1Horz: { shading: "#d9e2f3" },
    },
  });

  const mountTextPane = (spec: NamedStyleSpec, isNew: boolean): void => {
    held?.ctl.dispose();
    editorHost.replaceChildren();
    const s = sheet();
    const resolved = spec.basedOn || spec.id
      ? resolveStyle({ ...s, styles: spec.id && styleById(s, spec.id) ? s.styles : [...s.styles, { id: spec.id ?? "__cw_draft__", name: spec.name, type: spec.type, char: spec.char ?? {}, para: spec.para ?? {}, ...(spec.basedOn ? { basedOn: spec.basedOn } : {}) }] }, spec.id ?? "__cw_draft__")
      : { char: {}, para: {} };
    const ctl = mountTextStyleEditor(editorHost, {
      spec,
      resolvedChar: resolved.char,
      resolvedPara: resolved.para,
      basedOnOptions: basedOnOptions(spec.type, spec.id ?? null),
      isNew,
    });
    held = { kind: "text", ctl };
    ctl.onChange(refreshPreview);
    refreshPreview();
    syncFooter();
  };

  const mountListPane = (def: ListDefinition, isNew: boolean): void => {
    held?.ctl.dispose();
    editorHost.replaceChildren();
    const ctl = mountListStyleEditor(editorHost, { def, isNew });
    held = { kind: "list", ctl };
    ctl.onChange(refreshPreview);
    refreshPreview();
    syncFooter();
  };

  const mountTablePane = (style: TableStyle, isNew: boolean): void => {
    held?.ctl.dispose();
    editorHost.replaceChildren();
    const ctl = mountTableStyleEditor(editorHost, { style, isNew });
    held = { kind: "table", ctl };
    ctl.onChange(refreshPreview);
    refreshPreview();
    syncFooter();
  };

  const selectStyle = (kind: StyleKind, id: string): void => {
    selKind = kind;
    selId = id;
    if (kind === "paragraph" || kind === "character") {
      const st = styleById(sheet(), id);
      if (!st) return;
      const spec: NamedStyleSpec = { id, name: st.name, type: styleType(st), char: { ...st.char }, para: { ...st.para } };
      if (st.basedOn) spec.basedOn = st.basedOn;
      mountTextPane(spec, false);
    } else if (kind === "list") {
      const def = lists()[id];
      if (def) mountListPane(def, false);
    } else if (kind === "table") {
      const st = tableStyles()[id];
      if (st) mountTablePane(st, false);
    }
    rebuildList();
  };

  const newStyle = (kind: StyleKind, seedFrom?: NamedStyle): void => {
    selKind = kind;
    selId = null;
    if (kind === "paragraph" || kind === "character") {
      const base = kind === "paragraph" ? sheet().defaultStyleId : undefined;
      const spec: NamedStyleSpec = seedFrom
        ? { name: `${seedFrom.name} Copy`, type: kind, char: { ...seedFrom.char }, para: { ...seedFrom.para }, ...(seedFrom.basedOn ? { basedOn: seedFrom.basedOn } : {}) }
        : { name: kind === "character" ? "New Character Style" : "New Paragraph Style", type: kind, char: {}, para: {}, ...(base ? { basedOn: base } : {}) };
      mountTextPane(spec, true);
    } else if (kind === "list") {
      // Seed with a unique id so it doesn't clobber an existing definition.
      let n = 1;
      while (lists()[`custom-list-${n}`]) n++;
      mountListPane(bulletListDefinition(`custom-list-${n}`, "•"), true);
    } else if (kind === "table") {
      let n = 1;
      while (tableStyles()[`CustomTable${n}`]) n++;
      mountTablePane(defaultTableStyle(`CustomTable${n}`), true);
    }
    rebuildList();
  };

  // ---- List rail ------------------------------------------------------------
  const card = (kind: StyleKind, st: NamedStyle, isDefault: boolean): HTMLElement => {
    const c = el("button", "cw-sm-card");
    if (selKind === kind && selId === st.id) c.classList.add("sel");
    const nm = el("div", "nm");
    nm.append(document.createTextNode(st.name));
    if (isDefault) nm.append(el("span", "def", "default"));
    const sw = el("div", "sw");
    c.append(nm, sw);
    child.render(sw, { kind: "styleSample", styleId: st.id, sampleText: "AaBbCcDd" }, { fit: true, widthPx: 180, maxHeightPx: 22 });
    c.addEventListener("click", () => selectStyle(kind, st.id));
    c.addEventListener("contextmenu", (e) => {
      e.preventDefault();
      const entries: MenuEntry[] = [
        { kind: "item", label: "Modify", onClick: () => selectStyle(kind, st.id) },
        { kind: "item", label: "Duplicate", onClick: () => newStyle(kind, st) },
        { kind: "sep" },
        { kind: "item", label: "Set as default", disabled: kind !== "paragraph" || isDefault, onClick: () => { editor.dispatch(setDefaultStyle(st.id)); rebuildList(); } },
        { kind: "item", label: "Delete", danger: true, disabled: isDefault, onClick: () => deleteStyle(kind, st.id) },
      ];
      showContextMenu(e.clientX, e.clientY, entries);
    });
    return c;
  };

  const listCard = (def: ListDefinition): HTMLElement => {
    const c = el("button", "cw-sm-card");
    if (selKind === "list" && selId === def.id) c.classList.add("sel");
    const nm = el("div", "nm");
    nm.append(document.createTextNode(def.id));
    const sw = el("div", "sw");
    c.append(nm, sw);
    child.render(sw, { kind: "listSample", listDef: def, levels: 2 }, { fit: true, widthPx: 180, maxHeightPx: 34 });
    c.addEventListener("click", () => selectStyle("list", def.id));
    c.addEventListener("contextmenu", (e) => {
      e.preventDefault();
      showContextMenu(e.clientX, e.clientY, [
        { kind: "item", label: "Modify", onClick: () => selectStyle("list", def.id) },
        { kind: "item", label: "Delete", danger: true, onClick: () => deleteStyle("list", def.id) },
      ]);
    });
    return c;
  };

  const tableCard = (st: TableStyle): HTMLElement => {
    const c = el("button", "cw-sm-card");
    if (selKind === "table" && selId === st.id) c.classList.add("sel");
    const nm = el("div", "nm");
    nm.append(document.createTextNode(st.name));
    const sw = el("div", "sw");
    c.append(nm, sw);
    child.render(sw, { kind: "tableSample", tableStyle: st, rows: 3, cols: 3 }, { fit: true, widthPx: 180, maxHeightPx: 46 });
    c.addEventListener("click", () => selectStyle("table", st.id));
    c.addEventListener("contextmenu", (e) => {
      e.preventDefault();
      showContextMenu(e.clientX, e.clientY, [
        { kind: "item", label: "Modify", onClick: () => selectStyle("table", st.id) },
        { kind: "item", label: "Delete", danger: true, onClick: () => deleteStyle("table", st.id) },
      ]);
    });
    return c;
  };

  function rebuildList(): void {
    const s = sheet();
    listHost.replaceChildren();
    for (const kind of ["paragraph", "character", "list", "table"] as StyleKind[]) {
      const sec = el("div", "cw-sm-sec");
      sec.append(el("h4", undefined, KIND_LABELS[kind]));
      if (kind === "paragraph" || kind === "character") {
        const rows = s.styles.filter((x) => styleType(x) === kind);
        if (rows.length === 0) sec.append(el("div", "cw-sm-empty", "None yet"));
        for (const st of rows) sec.append(card(kind, st, s.defaultStyleId === st.id));
      } else if (kind === "list") {
        const defs = Object.values(lists());
        if (defs.length === 0) sec.append(el("div", "cw-sm-empty", "None yet"));
        for (const def of defs) sec.append(listCard(def));
      } else {
        const defs = Object.values(tableStyles());
        if (defs.length === 0) sec.append(el("div", "cw-sm-empty", "None yet"));
        for (const st of defs) sec.append(tableCard(st));
      }
      listHost.append(sec);
    }
  }

  // ---- Actions --------------------------------------------------------------
  const deriveId = (spec: NamedStyleSpec): string => spec.id ?? spec.name.trim().replace(/\s+/g, "");

  const apply = (): void => {
    if (!held) return;
    if (held.kind === "list") {
      const def = held.ctl.read();
      editor.dispatch(upsertListDefinition(def));
      if (lists()[def.id]) selectStyle("list", def.id);
      else rebuildList();
      return;
    }
    if (held.kind === "table") {
      const st = held.ctl.read();
      editor.dispatch(upsertTableStyle(st));
      if (tableStyles()[st.id]) selectStyle("table", st.id);
      else rebuildList();
      return;
    }
    const spec = held.ctl.read();
    if (spec.name.trim().length === 0) { editor.focus(); return; }
    const id = deriveId(spec);
    editor.dispatch(upsertNamedStyle(spec));
    // The style now exists under `id` — reselect it (also flips new → editing).
    if (styleById(sheet(), id)) selectStyle(spec.type as StyleKind, id);
    else rebuildList();
  };

  const deleteStyle = (kind: StyleKind, id: string): void => {
    if (kind === "paragraph" || kind === "character") editor.dispatch(deleteNamedStyle(id));
    else if (kind === "list") editor.dispatch(deleteListDefinition(id));
    else if (kind === "table") editor.dispatch(deleteTableStyle(id));
    if (selId === id) selectStyle("paragraph", sheet().defaultStyleId);
    else rebuildList();
  };

  function syncFooter(): void {
    const isDefault = selId !== null && sheet().defaultStyleId === selId;
    delBtn.disabled = selId === null || isDefault;
    const dupes = mergeDuplicateNamedStyles(sheet()).remap.size;
    mergeBtn.disabled = dupes === 0;
    mergeBtn.textContent = dupes > 0 ? `Merge duplicates (${dupes})` : "Merge duplicates";
  }

  newBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    const r = newBtn.getBoundingClientRect();
    const entries: MenuEntry[] = [
      { kind: "item", label: "Paragraph style", onClick: () => newStyle("paragraph") },
      { kind: "item", label: "Character style", onClick: () => newStyle("character") },
      { kind: "item", label: "List style", onClick: () => newStyle("list") },
      { kind: "item", label: "Table style", onClick: () => newStyle("table") },
    ];
    showContextMenu(r.left, r.bottom, entries);
  });
  applyBtn.addEventListener("click", apply);
  delBtn.addEventListener("click", () => { if (selId) deleteStyle(selKind, selId); });
  mergeBtn.addEventListener("click", () => {
    editor.dispatch(mergeDuplicateStyles());
    // The collapsed-into style may be the one selected — fall back to default.
    if (selId && !styleById(sheet(), selId)) selectStyle("paragraph", sheet().defaultStyleId);
    else { rebuildList(); syncFooter(); }
  });

  // ---- Lifecycle ------------------------------------------------------------
  const ac = new AbortController();
  const handle: StyleManagerHandle = {
    close(): void {
      held?.ctl.dispose();
      child.destroy();
      backdrop.remove();
      ac.abort();
      opts.onClose?.();
    },
    refresh(): void { rebuildList(); syncFooter(); },
  };
  window.addEventListener("keydown", (ev: KeyboardEvent) => {
    if (ev.key === "Escape") { ev.preventDefault(); ev.stopPropagation(); handle.close(); }
  }, { capture: true, signal: ac.signal });
  makeFloatingDialog({ backdrop, modal, handle: head, signal: ac.signal, noDrag: ".cw-sm-x" });
  xBtn.addEventListener("click", () => handle.close());
  closeBtn.addEventListener("click", () => handle.close());

  // ---- Initial selection ----------------------------------------------------
  rebuildList();
  const init = opts.initialSelection;
  const exists =
    init &&
    ((init.kind === "paragraph" || init.kind === "character") ? !!styleById(sheet(), init.id)
      : init.kind === "list" ? !!lists()[init.id]
      : init.kind === "table" ? !!tableStyles()[init.id]
      : false);
  if (init && exists) selectStyle(init.kind, init.id);
  else selectStyle("paragraph", sheet().defaultStyleId);

  return handle;
}
