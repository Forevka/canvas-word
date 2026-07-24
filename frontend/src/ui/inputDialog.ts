// A small, styled single-line input dialog — the real-UI replacement for the
// native prompt() calls that used to break the editor illusion (critique 2.8 /
// B5): bookmark add & rename, drop-down list items, etc. Built on createDialogShell
// so it inherits the surface manager's z-order + single-Escape (row 5) and the
// shared dialog look. Supports live validation (e.g. the OOXML bookmark-name
// rules) that a prompt() can neither enforce nor explain.

import { createDialogShell } from "./dialogShell";

const CSS = `
.cw-input-backdrop{position:fixed;inset:0;z-index:1100;background:rgba(20,22,26,.32);}
.cw-input-modal{position:fixed;min-width:320px;max-width:min(92vw,420px);background:#fff;border:1px solid #d0d5dd;
  border-radius:10px;box-shadow:0 12px 40px rgba(0,0,0,.22);font:13px/1.4 system-ui,Segoe UI,Arial,sans-serif;color:#1a1a2e;
  display:flex;flex-direction:column;overflow:hidden;}
.cw-input-head{display:flex;align-items:center;justify-content:space-between;padding:12px 14px 8px;cursor:move;}
.cw-input-head h2{margin:0;font-size:14px;font-weight:600;color:#201f1e;}
.cw-input-x{border:none;background:transparent;font-size:20px;line-height:1;cursor:pointer;color:#605e5c;padding:0 2px;}
.cw-input-x:hover{color:#201f1e;}
.cw-input-body{padding:2px 14px 6px;display:flex;flex-direction:column;gap:5px;}
.cw-input-body label{font-size:12px;color:#605e5c;}
.cw-input-body input{height:30px;border:1px solid #c8c6c4;border-radius:6px;padding:0 9px;font:inherit;font-size:13px;color:#201f1e;}
.cw-input-body input:focus{outline:none;border-color:#2b579a;box-shadow:0 0 0 1px #2b579a;}
.cw-input-body input.invalid{border-color:#a4262c;box-shadow:0 0 0 1px #a4262c;}
.cw-input-err{min-height:15px;font-size:11.5px;color:#a4262c;}
.cw-input-foot{display:flex;justify-content:flex-end;gap:8px;padding:8px 14px 12px;}
.cw-input-foot button{height:30px;padding:0 14px;border-radius:6px;font:inherit;font-size:13px;font-weight:600;cursor:pointer;border:1px solid #c8c6c4;background:#fff;color:#3c4043;}
.cw-input-foot button.primary{background:#2b579a;border-color:#2b579a;color:#fff;}
.cw-input-foot button.primary:disabled{opacity:.5;cursor:default;}
:root[data-theme="dark"] .cw-input-modal{background:#26282c;color:#e6e6e6;border-color:#4a4a4a;}
:root[data-theme="dark"] .cw-input-head h2{color:#f0f0f0;}
:root[data-theme="dark"] .cw-input-body label{color:#9aa0a6;}
:root[data-theme="dark"] .cw-input-body input{background:#1e1f22;color:#e6e6e6;border-color:#4a4a4a;}
:root[data-theme="dark"] .cw-input-foot button{background:#2a2c30;color:#e6e6e6;border-color:#4a4a4a;}
:root[data-theme="dark"] .cw-input-foot button.primary{background:#3d6ba5;border-color:#3d6ba5;color:#fff;}
`;

export interface InputDialogOptions {
  title: string;
  label: string;
  initial?: string;
  placeholder?: string;
  okLabel?: string;
  /** Return an error message to block submission, or null when the value is OK. */
  validate?: (value: string) => string | null;
  /** Called with the (validated) value on OK / Enter. */
  onSubmit: (value: string) => void;
}

/** Show the input dialog. Non-blocking (unlike prompt): the caller supplies an
 *  onSubmit callback. Returns the DialogShell in case the caller wants to close
 *  it programmatically. */
export function showInputDialog(o: InputDialogOptions): void {
  const shell = createDialogShell({ prefix: "cw-input", cssId: "cw-input-styles", css: CSS, title: o.title });

  const label = document.createElement("label");
  label.textContent = o.label;
  const input = document.createElement("input");
  input.type = "text";
  input.value = o.initial ?? "";
  if (o.placeholder) input.placeholder = o.placeholder;
  const err = document.createElement("div");
  err.className = "cw-input-err";
  shell.body.append(label, input, err);

  const cancel = document.createElement("button");
  cancel.textContent = "Cancel";
  const ok = document.createElement("button");
  ok.className = "primary";
  ok.textContent = o.okLabel ?? "OK";
  shell.foot.append(cancel, ok);

  const validateNow = (): string | null => {
    const msg = o.validate ? o.validate(input.value) : null;
    err.textContent = msg ?? "";
    input.classList.toggle("invalid", !!msg);
    ok.disabled = !!msg;
    return msg;
  };
  const submit = (): void => {
    if (validateNow() !== null) return;
    const value = input.value;
    shell.close();
    o.onSubmit(value);
  };
  input.addEventListener("input", validateNow);
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); submit(); }
    e.stopPropagation(); // don't leak keys to the editor; Escape is the shell's job
  });
  cancel.addEventListener("click", () => shell.close());
  ok.addEventListener("click", submit);

  validateNow();
  setTimeout(() => { input.focus(); input.select(); }, 0);
}
