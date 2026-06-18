// Content-control inspector — a modal popup that shows an SDT's properties and
// an EDITABLE preview of its wrapped content. Editing + Save writes the content
// back to the document (index.ts supplies the onSave handler). Useful for QA'ing
// auto-generated sections that get wrapped in content controls.
//
// Pure presentation: index.ts gathers the data (properties + rendered HTML of
// the control's content) and handles persistence.

import type { Block, SdtProps } from "@cw/shared";
import { injectCssOnce } from "./styles";
import { makeFloatingDialog } from "./floatingDialog";

export interface SdtInspectorData {
  id: string;
  props: SdtProps;
  /** Plain-text content (for the empty check + copy + char count). */
  text: string;
  /** How many paragraphs the control spans (block-level controls wrap many). */
  paragraphCount: number;
  charCount: number;
}

/** Live editor over the control's content, mounted into the preview host. */
export interface SdtInspectorEditor {
  getBlocks(): Block[];
  /** Insert an image (bytes registered in the document's shared media store). */
  insertImage(bytes: Uint8Array, mime: string, opts?: { widthPx?: number; heightPx?: number }): Promise<void>;
}

export interface SdtInspectorOptions {
  /** When false the content is shown read-only (locked controls, or hosts the
   *  replace command can't restructure). */
  editable?: boolean;
  /** Paint the control's content into the preview host (read-only path) — a real
   *  canvas render via a child document, not HTML. */
  renderPreview?: (host: HTMLElement) => void;
  /** Mount a canvas-native editor into the preview host (editable path). Save
   *  reads getBlocks() and hands them to onSave. */
  mountEditor?: (host: HTMLElement) => SdtInspectorEditor;
  /** Commit edited content (the child editor's blocks). Returns true on success. */
  onSave?: (blocks: Block[]) => boolean;
  /** Show an "Insert image" affordance (block-level rich-text controls only — the
   *  inline/cell commit paths are text-only). */
  allowImages?: boolean;
  /** Called when the inspector closes — tear down the child document. */
  onClose?: () => void;
}

export interface SdtInspectorHandle {
  close(): void;
}

const SDT_TYPE_LABEL: Record<SdtProps["type"], string> = {
  richText: "Rich Text",
  plainText: "Plain Text",
  checkbox: "Check Box",
  dropDown: "Drop-Down List",
  comboBox: "Combo Box",
  date: "Date Picker",
};

const SDT_CSS = `
.cw-sdt-backdrop{position:fixed;inset:0;z-index:1100;background:rgba(20,22,26,.38);
  display:flex;align-items:center;justify-content:center;}
.cw-sdt-modal{width:min(960px,94vw);max-height:88vh;display:flex;flex-direction:column;
  background:#fff;border-radius:10px;box-shadow:0 18px 56px rgba(0,0,0,.34);
  font:13px/1.5 Arial,sans-serif;color:#202124;overflow:hidden;}
.cw-sdt-head{display:flex;align-items:center;gap:10px;padding:13px 16px;border-bottom:1px solid #e6e8eb;}
.cw-sdt-head h2{margin:0;font-size:15px;font-weight:600;flex:1 1 auto;white-space:nowrap;
  overflow:hidden;text-overflow:ellipsis;}
.cw-sdt-badge{font-size:11px;font-weight:600;color:#0b57d0;background:#e8f0fe;border-radius:10px;padding:2px 9px;}
.cw-sdt-x{border:none;background:transparent;font-size:20px;line-height:1;color:#5f6368;cursor:pointer;
  width:28px;height:28px;border-radius:6px;}
.cw-sdt-x:hover{background:#e8eaed;}
/* Side-by-side: properties on the left, content on the right. */
.cw-sdt-body{display:flex;min-height:0;flex:1 1 auto;}
.cw-sdt-left{flex:0 0 300px;padding:12px 16px;overflow:auto;border-right:1px solid #e6e8eb;background:#fbfbfc;}
.cw-sdt-right{flex:1 1 auto;display:flex;flex-direction:column;padding:12px 16px;min-width:0;}
.cw-sdt-section-title{font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:#80868b;
  margin:0 0 8px;display:flex;align-items:center;gap:8px;}
.cw-sdt-section-title .hint{font-weight:400;text-transform:none;letter-spacing:0;color:#9aa0a6;}
.cw-sdt-props{display:grid;grid-template-columns:auto 1fr;gap:6px 10px;}
.cw-sdt-props dt{color:#5f6368;}
.cw-sdt-props dd{margin:0;word-break:break-word;}
.cw-sdt-props dd code{background:#eef0f2;border-radius:4px;padding:1px 5px;font-size:11px;
  font-family:Consolas,monospace;word-break:break-all;}
.cw-sdt-pill{display:inline-block;background:#eef0f2;border-radius:10px;padding:1px 8px;margin:0 4px 4px 0;}
.cw-sdt-preview{flex:1 1 auto;border:1px solid #dadce0;border-radius:8px;padding:14px 16px;background:#fff;
  min-height:120px;overflow:auto;}
.cw-sdt-preview.editable:focus{border-color:#1a73e8;box-shadow:0 0 0 2px rgba(26,115,232,.18);outline:none;}
.cw-sdt-preview p{margin:0 0 .5em;}
.cw-sdt-preview img{max-width:100%;height:auto;}
.cw-sdt-empty{color:#80868b;font-style:italic;}
.cw-sdt-foot{display:flex;align-items:center;gap:8px;padding:11px 16px;border-top:1px solid #e6e8eb;}
.cw-sdt-foot .spacer{flex:1 1 auto;}
.cw-sdt-btn{height:30px;padding:0 14px;border:1px solid #d0d4d9;border-radius:6px;background:#fff;
  cursor:pointer;font-size:13px;color:#3c4043;}
.cw-sdt-btn:hover{background:#f1f3f4;}
.cw-sdt-btn.primary{border-color:#1a73e8;background:#1a73e8;color:#fff;}
.cw-sdt-btn.primary:hover{background:#1864cc;}
.cw-sdt-btn:disabled{opacity:.5;cursor:default;}`;

export function showSdtInspector(
  data: SdtInspectorData,
  opts: SdtInspectorOptions = {},
): SdtInspectorHandle {
  injectCssOnce("cw-sdt-styles", SDT_CSS);
  const { props } = data;
  const editable =
    opts.editable === true && typeof opts.onSave === "function" && typeof opts.mountEditor === "function";
  let childEditor: SdtInspectorEditor | null = null;
  const title = props.alias?.trim() || SDT_TYPE_LABEL[props.type];

  const backdrop = document.createElement("div");
  backdrop.className = "cw-sdt-backdrop";
  const modal = document.createElement("div");
  modal.className = "cw-sdt-modal";
  modal.addEventListener("mousedown", (e) => e.stopPropagation());

  // Header
  const head = document.createElement("div");
  head.className = "cw-sdt-head";
  const h2 = document.createElement("h2");
  h2.textContent = title;
  const badge = document.createElement("span");
  badge.className = "cw-sdt-badge";
  badge.textContent = SDT_TYPE_LABEL[props.type];
  const xBtn = document.createElement("button");
  xBtn.className = "cw-sdt-x";
  xBtn.textContent = "×";
  xBtn.title = "Close (Esc)";
  head.append(h2, badge, xBtn);

  // Body: two columns — properties (left) and content (right).
  const body = document.createElement("div");
  body.className = "cw-sdt-body";
  const left = document.createElement("div");
  left.className = "cw-sdt-left";
  const right = document.createElement("div");
  right.className = "cw-sdt-right";

  const propsTitle = document.createElement("div");
  propsTitle.className = "cw-sdt-section-title";
  propsTitle.textContent = "Properties";
  const dl = document.createElement("dl");
  dl.className = "cw-sdt-props";
  const addProp = (label: string, valueEl: Node | string): void => {
    const dt = document.createElement("dt");
    dt.textContent = label;
    const dd = document.createElement("dd");
    if (typeof valueEl === "string") dd.textContent = valueEl;
    else dd.appendChild(valueEl);
    dl.append(dt, dd);
  };
  const codeEl = (s: string): HTMLElement => {
    const c = document.createElement("code");
    c.textContent = s;
    return c;
  };

  addProp("Type", SDT_TYPE_LABEL[props.type]);
  addProp("Title (alias)", props.alias ? codeEl(props.alias) : italic("—"));
  addProp("Tag", props.tag ? codeEl(props.tag) : italic("—"));
  if (props.type === "checkbox") addProp("Checked", props.checked ? "Yes" : "No");
  if (props.type === "date") addProp("Date format", props.dateFormat ? codeEl(props.dateFormat) : italic("default"));
  if (props.listItems && props.listItems.length > 0) {
    const wrap = document.createElement("div");
    for (const li of props.listItems) {
      const pill = document.createElement("span");
      pill.className = "cw-sdt-pill";
      pill.textContent = li.display + (li.value !== li.display ? ` (${li.value})` : "");
      wrap.appendChild(pill);
    }
    addProp(`List items (${props.listItems.length})`, wrap);
  }
  const locks: string[] = [];
  if (props.lockContent) locks.push("Content can't be edited");
  if (props.lockControl) locks.push("Control can't be deleted");
  addProp("Locks", locks.length ? locks.join("; ") : italic("none"));
  addProp("Placeholder", props.placeholder ? "Showing placeholder text" : italic("no"));
  addProp("Spans", `${data.paragraphCount} paragraph${data.paragraphCount === 1 ? "" : "s"} · ${data.charCount} chars`);
  addProp("Control id", codeEl(data.id));

  const contentTitle = document.createElement("div");
  contentTitle.className = "cw-sdt-section-title";
  contentTitle.textContent = "Content";
  const hint = document.createElement("span");
  hint.className = "hint";
  hint.textContent = editable ? "— editable; Save to apply" : "— read-only";
  contentTitle.appendChild(hint);

  const preview = document.createElement("div");
  preview.className = "cw-sdt-preview";
  if (editable) preview.classList.add("editable");
  // Content is painted (read-only) or an editor mounted (editable) AFTER the modal
  // is in the DOM, so the host has a measurable width — see the mount step below.

  left.append(propsTitle, dl);
  right.append(contentTitle, preview);
  body.append(left, right);

  // Footer
  const foot = document.createElement("div");
  foot.className = "cw-sdt-foot";
  const copyBtn = mkBtn("Copy Text", false);
  copyBtn.addEventListener("click", () => {
    void navigator.clipboard?.writeText(data.text).then(
      () => (copyBtn.textContent = "Copied ✓"),
      () => (copyBtn.textContent = "Copy failed"),
    );
    setTimeout(() => (copyBtn.textContent = "Copy Text"), 1400);
  });
  const spacer = document.createElement("div");
  spacer.className = "spacer";
  foot.append(copyBtn, spacer);

  if (editable && opts.allowImages) {
    const imgBtn = mkBtn("Insert Image…", false);
    imgBtn.addEventListener("click", () => {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = "image/*";
      input.addEventListener("change", () => {
        const file = input.files?.[0];
        if (!file || !childEditor) return;
        void file.arrayBuffer().then((buf) => childEditor!.insertImage(new Uint8Array(buf), file.type || "image/png"));
      });
      input.click();
    });
    foot.append(imgBtn);
  }
  if (editable) {
    const saveBtn = mkBtn("Save Changes", true);
    saveBtn.addEventListener("click", () => {
      const ok = childEditor ? opts.onSave!(childEditor.getBlocks()) : false;
      if (ok) {
        saveBtn.textContent = "Saved ✓";
        setTimeout(() => handle.close(), 450);
      } else {
        saveBtn.textContent = "Couldn't save";
        setTimeout(() => (saveBtn.textContent = "Save Changes"), 1600);
      }
    });
    const closeBtn = mkBtn("Cancel", false);
    closeBtn.addEventListener("click", () => handle.close());
    foot.append(closeBtn, saveBtn);
  } else {
    const closeBtn = mkBtn("Close", true);
    closeBtn.addEventListener("click", () => handle.close());
    foot.append(closeBtn);
  }

  modal.append(head, body, foot);
  backdrop.appendChild(modal);
  document.body.appendChild(backdrop);

  // Now that the preview host has a measurable width, paint the content (read-only)
  // or mount the canvas-native editor (editable).
  if (data.text.trim() === "" && !editable) {
    preview.innerHTML = `<span class="cw-sdt-empty">(empty)</span>`;
  } else if (editable && opts.mountEditor) {
    childEditor = opts.mountEditor(preview);
  } else {
    opts.renderPreview?.(preview);
  }

  const ac = new AbortController();
  const handle: SdtInspectorHandle = {
    close(): void {
      backdrop.remove();
      ac.abort(); // detaches every listener registered with ac.signal
      opts.onClose?.(); // tear down the child document
    },
  };
  window.addEventListener(
    "keydown",
    (ev: KeyboardEvent) => {
      if (ev.key === "Escape") {
        ev.preventDefault();
        ev.stopPropagation();
        handle.close();
      }
    },
    { capture: true, signal: ac.signal },
  );
  makeFloatingDialog({ backdrop, modal, handle: head, signal: ac.signal, noDrag: ".cw-sdt-x" });
  xBtn.addEventListener("click", () => handle.close());
  return handle;
}

function mkBtn(label: string, primary: boolean): HTMLButtonElement {
  const b = document.createElement("button");
  b.className = primary ? "cw-sdt-btn primary" : "cw-sdt-btn";
  b.textContent = label;
  return b;
}

function italic(text: string): HTMLElement {
  const i = document.createElement("i");
  i.style.color = "#9aa0a6";
  i.textContent = text;
  return i;
}
