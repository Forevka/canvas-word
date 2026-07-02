// TOC properties — a modal to review and edit a table-of-contents field's OOXML
// switches (\o level range, \u outline levels, \h hyperlinks, \n hidden page
// numbers, \p separator, \z web-only hide). On Apply the caller turns the edited
// TocSwitches into ops (setTocSwitchesCmd): persist the instruction + regenerate
// the entries. Pure presentation, modeled on the field constructor modal.

import type { TocSwitches } from "@cw/shared";
import { createDialogShell } from "./dialogShell";

export interface TocPropertiesOptions {
  /** Current switches, parsed from the TOC field's instruction. */
  initial: TocSwitches;
  onApply: (sw: TocSwitches) => void;
}

export interface TocPropertiesHandle {
  close(): void;
}

const CSS = `
.cw-toc-backdrop{position:fixed;inset:0;z-index:1100;background:rgba(20,22,26,.38);display:flex;align-items:center;justify-content:center;}
.cw-toc-modal{width:min(560px,94vw);max-height:88vh;display:flex;flex-direction:column;background:#fff;border-radius:10px;
  box-shadow:0 18px 56px rgba(0,0,0,.34);font:13px/1.5 Arial,sans-serif;color:#202124;overflow:hidden;}
.cw-toc-head{display:flex;align-items:center;gap:10px;padding:13px 16px;border-bottom:1px solid #e6e8eb;}
.cw-toc-head h2{margin:0;font-size:15px;font-weight:600;flex:1 1 auto;}
.cw-toc-x{border:none;background:transparent;font-size:20px;line-height:1;color:#5f6368;cursor:pointer;width:28px;height:28px;border-radius:6px;}
.cw-toc-x:hover{background:#e8eaed;}
.cw-toc-body{padding:14px 16px;overflow:auto;display:flex;flex-direction:column;gap:14px;}
.cw-toc-field{display:flex;flex-direction:column;gap:4px;}
.cw-toc-field>label{font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:#80868b;}
.cw-toc-check{display:flex;align-items:center;gap:8px;font-size:13px;color:#3c4043;}
.cw-toc-check input{width:16px;height:16px;}
.cw-toc-range{display:flex;align-items:center;gap:8px;}
.cw-toc-range input[type=number]{width:64px;}
.cw-toc-range.dim{opacity:.45;}
.cw-toc-body input[type=number],.cw-toc-body input[type=text]{height:30px;border:1px solid #d0d4d9;border-radius:6px;padding:0 8px;font:13px Arial,sans-serif;}
.cw-toc-instr{font:12px/1.4 Consolas,monospace;color:#5f6368;background:#f6f7f9;border:1px solid #e6e8eb;border-radius:6px;padding:8px 10px;word-break:break-all;}
.cw-toc-foot{display:flex;align-items:center;gap:8px;padding:11px 16px;border-top:1px solid #e6e8eb;}
.cw-toc-foot .spacer{flex:1 1 auto;}
.cw-toc-btn{height:30px;padding:0 14px;border:1px solid #d0d4d9;border-radius:6px;background:#fff;cursor:pointer;font-size:13px;color:#3c4043;}
.cw-toc-btn:hover{background:#f1f3f4;}
.cw-toc-btn.primary{border-color:#1a73e8;background:#1a73e8;color:#fff;}
.cw-toc-btn.primary:hover{background:#1864cc;}`;

const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string, text?: string): HTMLElementTagNameMap[K] => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
};

const checkbox = (label: string, checked: boolean): { row: HTMLElement; input: HTMLInputElement } => {
  const row = el("label", "cw-toc-check");
  const input = el("input");
  input.type = "checkbox";
  input.checked = checked;
  row.append(input, document.createTextNode(label));
  return { row, input };
};

export function showTocProperties(opts: TocPropertiesOptions): TocPropertiesHandle {
  const init = opts.initial;

  const shell = createDialogShell({
    prefix: "cw-toc",
    cssId: "cw-toc-styles",
    css: CSS,
    title: "Table of contents — field options",
  });
  const { body, foot } = shell;

  // \o "from-to" — heading level range
  const oField = el("div", "cw-toc-field");
  const oToggle = checkbox("Build from heading levels (\\o)", !!init.outlineRange);
  const oRange = el("div", "cw-toc-range");
  const fromIn = el("input"); fromIn.type = "number"; fromIn.min = "1"; fromIn.max = "9"; fromIn.value = String(init.outlineRange?.from ?? 1);
  const toIn = el("input"); toIn.type = "number"; toIn.min = "1"; toIn.max = "9"; toIn.value = String(init.outlineRange?.to ?? 3);
  oRange.append(el("span", undefined, "Levels"), fromIn, el("span", undefined, "to"), toIn);
  oField.append(el("label", undefined, "Heading levels"), oToggle.row, oRange);
  body.append(oField);

  // \u, \h, \z
  const uChk = checkbox("Include paragraphs with a direct outline level (\\u)", init.useOutlineLevels);
  const hChk = checkbox("Make entries hyperlinks to their heading (\\h)", init.hyperlinks);
  const zChk = checkbox("Hide leader & page numbers in web view (\\z)", init.hideInWeb);
  const flags = el("div", "cw-toc-field");
  flags.append(el("label", undefined, "Options"), uChk.row, hChk.row, zChk.row);
  body.append(flags);

  // \p separator
  const sepField = el("div", "cw-toc-field");
  const sepIn = el("input"); sepIn.type = "text"; sepIn.placeholder = "tab (default)"; sepIn.value = init.separator ?? "";
  sepField.append(el("label", undefined, "Label / page-number separator (\\p)"), sepIn);
  body.append(sepField);

  // Live instruction preview
  const instrField = el("div", "cw-toc-field");
  const instr = el("div", "cw-toc-instr");
  instrField.append(el("label", undefined, "Field instruction"), instr);
  body.append(instrField);

  const cancel = el("button", "cw-toc-btn", "Cancel");
  const apply = el("button", "cw-toc-btn primary", "Apply");
  foot.append(el("div", "spacer"), cancel, apply);

  const read = (): TocSwitches => {
    const sw: TocSwitches = {
      useOutlineLevels: uChk.input.checked,
      hyperlinks: hChk.input.checked,
      hideInWeb: zChk.input.checked,
    };
    if (oToggle.input.checked) {
      const from = Math.max(1, Math.min(9, Number(fromIn.value) || 1));
      const to = Math.max(from, Math.min(9, Number(toIn.value) || from));
      sw.outlineRange = { from, to };
    }
    if (init.customStyles) sw.customStyles = init.customStyles; // preserve \t mappings we don't edit here
    if (init.hidePageNumberRange) sw.hidePageNumberRange = init.hidePageNumberRange;
    if (sepIn.value.length > 0) sw.separator = sepIn.value;
    return sw;
  };

  const refresh = (): void => {
    oRange.classList.toggle("dim", !oToggle.input.checked);
    fromIn.disabled = toIn.disabled = !oToggle.input.checked;
    // Inline-built instruction string (mirrors shared buildTocInstruction).
    const sw = read();
    const parts = ["TOC"];
    if (sw.outlineRange) parts.push(`\\o "${sw.outlineRange.from}-${sw.outlineRange.to}"`);
    if (sw.useOutlineLevels) parts.push("\\u");
    if (sw.hyperlinks) parts.push("\\h");
    if (sw.separator !== undefined) parts.push(`\\p "${sw.separator}"`);
    if (sw.hideInWeb) parts.push("\\z");
    instr.textContent = ` ${parts.join(" ")} `;
  };
  for (const c of [oToggle.input, uChk.input, hChk.input, zChk.input]) c.addEventListener("change", refresh);
  for (const i of [fromIn, toIn, sepIn]) i.addEventListener("input", refresh);

  const handle: TocPropertiesHandle = { close: shell.close };
  cancel.addEventListener("click", shell.close);
  apply.addEventListener("click", () => { opts.onApply(read()); shell.close(); });

  refresh();
  return handle;
}
