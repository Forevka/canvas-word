// Field constructor — a modal to insert or edit a built-in field (PAGE, NUMPAGES,
// DATE, TIME, IF). The left column picks the field type + its arguments/switches;
// the right column shows a live preview of the result. On Apply, the caller turns
// the FieldSpec into ops (insertFieldCmd / editFieldCmd). Pure presentation.

import type { CharStyle, FieldSpec, IfOp, PageNumFmt, Run } from "@cw/shared";
import { evaluateIf, formatFieldDate } from "@cw/shared";
import { createDialogShell } from "./dialogShell";

export interface FieldConstructorOptions {
  /** Pre-fill for editing an existing field; omit to insert a new one. */
  initial?: FieldSpec;
  /** Bookmark names in the document — the targets a REF/PAGEREF can point at. When
   *  empty, the cross-reference types are offered but explain that none exist yet. */
  bookmarks?: string[];
  /** Base run style for the result preview (and IF result text). */
  baseStyle: CharStyle;
  /** "now" for the DATE/TIME preview. */
  now?: Date;
  onApply: (spec: FieldSpec) => void;
  /** Paint the result preview into `host` in the document's real font (via a child
   *  document) instead of plain text. Reuses `host` across refreshes. */
  renderResult?: (host: HTMLElement, runs: Run[]) => void;
  /** Called when the dialog closes — tear down the child document. */
  onClose?: () => void;
}

export interface FieldConstructorHandle {
  close(): void;
}

type FieldType = FieldSpec["type"];

const TYPES: { value: FieldType; label: string }[] = [
  { value: "PAGE", label: "Page number (PAGE)" },
  { value: "NUMPAGES", label: "Total pages (NUMPAGES)" },
  { value: "DATE", label: "Date (DATE)" },
  { value: "TIME", label: "Time (TIME)" },
  { value: "IF", label: "Conditional (IF)" },
  { value: "REF", label: "Cross-reference — text (REF)" },
  { value: "PAGEREF", label: "Cross-reference — page (PAGEREF)" },
];
const NUM_FMTS: { value: PageNumFmt; label: string }[] = [
  { value: "arabic", label: "1, 2, 3" },
  { value: "roman", label: "i, ii, iii" },
  { value: "Roman", label: "I, II, III" },
  { value: "alpha", label: "a, b, c" },
  { value: "Alpha", label: "A, B, C" },
];
const IF_OPS: IfOp[] = ["=", "<>", "<", ">", "<=", ">="];
const DATE_PRESETS = ["M/d/yyyy", "MMMM d, yyyy", "dddd, MMMM d, yyyy", "yyyy-MM-dd"];
const TIME_PRESETS = ["h:mm AM/PM", "HH:mm", "h:mm:ss AM/PM"];

const CSS = `
.cw-fc-backdrop{position:fixed;inset:0;z-index:1100;background:rgba(20,22,26,.38);display:flex;align-items:center;justify-content:center;}
.cw-fc-modal{width:min(720px,94vw);max-height:88vh;display:flex;flex-direction:column;background:#fff;border-radius:10px;
  box-shadow:0 18px 56px rgba(0,0,0,.34);font:13px/1.5 Arial,sans-serif;color:#202124;overflow:hidden;}
.cw-fc-head{display:flex;align-items:center;gap:10px;padding:13px 16px;border-bottom:1px solid #e6e8eb;}
.cw-fc-head h2{margin:0;font-size:15px;font-weight:600;flex:1 1 auto;}
.cw-fc-x{border:none;background:transparent;font-size:20px;line-height:1;color:#5f6368;cursor:pointer;width:28px;height:28px;border-radius:6px;}
.cw-fc-x:hover{background:#e8eaed;}
.cw-fc-body{display:flex;min-height:0;flex:1 1 auto;}
.cw-fc-left{flex:1 1 auto;padding:14px 16px;overflow:auto;display:flex;flex-direction:column;gap:12px;}
.cw-fc-right{flex:0 0 240px;border-left:1px solid #e6e8eb;background:#fbfbfc;padding:14px 16px;display:flex;flex-direction:column;gap:8px;}
.cw-fc-field{display:flex;flex-direction:column;gap:4px;}
.cw-fc-field label{font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:#80868b;}
.cw-fc-field input,.cw-fc-field select{height:30px;border:1px solid #d0d4d9;border-radius:6px;padding:0 8px;font:13px Arial,sans-serif;}
.cw-fc-row{display:flex;gap:8px;}.cw-fc-row>*{flex:1 1 auto;}
.cw-fc-presets{display:flex;flex-wrap:wrap;gap:6px;}
.cw-fc-chip{border:1px solid #d0d4d9;border-radius:14px;background:#fff;cursor:pointer;font-size:11px;padding:3px 9px;color:#3c4043;}
.cw-fc-chip:hover{background:#f1f3f4;}
.cw-fc-prev-title{font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:#80868b;}
.cw-fc-prev{border:1px solid #dadce0;border-radius:8px;padding:10px 12px;background:#fff;min-height:48px;font-size:14px;word-break:break-word;}
.cw-fc-note{font-size:11px;color:#9aa0a6;}
.cw-fc-foot{display:flex;align-items:center;gap:8px;padding:11px 16px;border-top:1px solid #e6e8eb;}
.cw-fc-foot .spacer{flex:1 1 auto;}
.cw-fc-btn{height:30px;padding:0 14px;border:1px solid #d0d4d9;border-radius:6px;background:#fff;cursor:pointer;font-size:13px;color:#3c4043;}
.cw-fc-btn:hover{background:#f1f3f4;}
.cw-fc-btn.primary{border-color:#1a73e8;background:#1a73e8;color:#fff;}
.cw-fc-btn.primary:hover{background:#1864cc;}
.cw-fc-btn:disabled{opacity:.5;cursor:not-allowed;}
.cw-fc-btn.primary:disabled:hover{background:#1a73e8;}`;

const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string, text?: string): HTMLElementTagNameMap[K] => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
};

const labelled = (label: string, control: HTMLElement): HTMLElement => {
  const wrap = el("div", "cw-fc-field");
  wrap.append(el("label", undefined, label), control);
  return wrap;
};

const option = (value: string, label: string): HTMLOptionElement => {
  const o = el("option");
  o.value = value;
  o.textContent = label;
  return o;
};

export function showFieldConstructor(opts: FieldConstructorOptions): FieldConstructorHandle {
  const now = opts.now ?? new Date();
  const editing = opts.initial !== undefined;

  const shell = createDialogShell({
    prefix: "cw-fc",
    cssId: "cw-fc-styles",
    css: CSS,
    title: editing ? "Edit field" : "Insert field",
    ...(opts.onClose ? { onClose: opts.onClose } : {}), // tear down the child document
  });
  const { body, foot } = shell;
  const left = el("div", "cw-fc-left");
  const right = el("div", "cw-fc-right");

  // Type picker
  const typeSel = el("select");
  for (const t of TYPES) typeSel.append(option(t.value, t.label));
  typeSel.value = opts.initial?.type ?? "PAGE";
  left.append(labelled("Field type", typeSel));

  const formHost = el("div", "cw-fc-field");
  left.append(formHost);

  const prevTitle = el("div", "cw-fc-prev-title", "Preview");
  const prev = el("div", "cw-fc-prev");
  const note = el("div", "cw-fc-note");
  right.append(prevTitle, prev, note);

  // --- per-type form builders. Each returns a readSpec() closure. -------------
  let readSpec: () => FieldSpec = () => ({ type: "PAGE", numFmt: "arabic" });

  const numFmtForm = (type: "PAGE" | "NUMPAGES"): (() => FieldSpec) => {
    const sel = el("select");
    for (const f of NUM_FMTS) sel.append(option(f.value, f.label));
    const init = opts.initial;
    sel.value = (init && (init.type === "PAGE" || init.type === "NUMPAGES") && init.numFmt) || "arabic";
    sel.addEventListener("change", refresh);
    formHost.append(labelled("Number format", sel));
    return () => ({ type, numFmt: sel.value as PageNumFmt });
  };

  const dateForm = (type: "DATE" | "TIME"): (() => FieldSpec) => {
    const input = el("input");
    input.type = "text";
    const init = opts.initial;
    input.value = (init && (init.type === "DATE" || init.type === "TIME") && init.format) || (type === "DATE" ? "M/d/yyyy" : "h:mm AM/PM");
    input.addEventListener("input", refresh);
    const presets = el("div", "cw-fc-presets");
    for (const p of type === "DATE" ? DATE_PRESETS : TIME_PRESETS) {
      const chip = el("button", "cw-fc-chip", p);
      chip.type = "button";
      chip.addEventListener("click", () => { input.value = p; refresh(); });
      presets.append(chip);
    }
    formHost.append(labelled("Format (Word \\@ picture)", input), presets);
    return () => ({ type, format: input.value });
  };

  const ifForm = (): (() => FieldSpec) => {
    const a = el("input"); a.type = "text"; a.placeholder = "e.g. 1";
    const opSel = el("select");
    for (const o of IF_OPS) opSel.append(option(o, o));
    const b = el("input"); b.type = "text"; b.placeholder = "e.g. 1";
    const t = el("input"); t.type = "text"; t.placeholder = "shown when true";
    const f = el("input"); f.type = "text"; f.placeholder = "shown when false";
    const init = opts.initial;
    if (init?.type === "IF") {
      a.value = init.operandA; opSel.value = init.op; b.value = init.operandB;
      t.value = init.trueRuns.map((r) => r.text).join("");
      f.value = init.falseRuns.map((r) => r.text).join("");
    }
    for (const inp of [a, opSel, b, t, f]) inp.addEventListener("input", refresh), inp.addEventListener("change", refresh);
    const cmp = el("div", "cw-fc-row");
    cmp.append(labelled("If", a), labelled("Is", opSel), labelled("Than", b));
    formHost.append(cmp, labelled("Then show", t), labelled("Otherwise show", f));
    return () => ({
      type: "IF", operandA: a.value, op: opSel.value as IfOp, operandB: b.value,
      trueRuns: [{ text: t.value, style: { ...opts.baseStyle } }],
      falseRuns: [{ text: f.value, style: { ...opts.baseStyle } }],
    });
  };

  const refForm = (type: "REF" | "PAGEREF"): (() => FieldSpec) => {
    const bookmarks = opts.bookmarks ?? [];
    const sel = el("select");
    for (const b of bookmarks) sel.append(option(b, b));
    const init = opts.initial;
    const prior = init && (init.type === "REF" || init.type === "PAGEREF") ? init.bookmark : undefined;
    if (prior !== undefined && bookmarks.includes(prior)) sel.value = prior;
    else if (prior) {
      // The field's target was deleted or renamed. Leaving the <select> on whatever
      // sorts first would silently retarget the field on Apply, so surface the miss
      // as a selected placeholder with an empty value — Apply stays disabled (the
      // no-target check in refresh) until a real bookmark is chosen.
      const missing = option("", `${prior} — missing`);
      sel.prepend(missing);
      sel.value = "";
    }
    sel.disabled = bookmarks.length === 0;
    sel.addEventListener("change", refresh);
    formHost.append(labelled("Bookmark", sel));
    if (bookmarks.length === 0) {
      const hint = el("div", "cw-fc-note", "No bookmarks in this document yet. Add one from Navigator ▸ Marks, then reference it here.");
      formHost.append(hint);
    }
    return () => ({ type, bookmark: sel.value });
  };

  function rebuildForm(): void {
    formHost.replaceChildren();
    const type = typeSel.value as FieldType;
    readSpec =
      type === "PAGE" || type === "NUMPAGES" ? numFmtForm(type)
      : type === "DATE" || type === "TIME" ? dateForm(type)
      : type === "REF" || type === "PAGEREF" ? refForm(type)
      : ifForm();
    refresh();
  }

  // The field result as styled runs. PAGE/NUMPAGES use sample values (the dialog
  // has no pagination context); DATE/TIME/IF are materialized from the spec.
  const resultRuns = (spec: FieldSpec): Run[] => {
    switch (spec.type) {
      case "PAGE":
        return [{ text: "1", style: { ...opts.baseStyle } }];
      case "NUMPAGES":
        return [{ text: "12", style: { ...opts.baseStyle } }];
      case "DATE":
      case "TIME":
        return [{ text: formatFieldDate(now, spec.format) || "(empty format)", style: { ...opts.baseStyle } }];
      case "REF":
        return [{ text: spec.bookmark || "(no bookmark)", style: { ...opts.baseStyle } }];
      case "PAGEREF":
        return [{ text: spec.bookmark ? "1" : "(no bookmark)", style: { ...opts.baseStyle } }];
      default: {
        const runs = evaluateIf(spec);
        return runs.length ? runs : [{ text: "(empty)", style: { ...opts.baseStyle } }];
      }
    }
  };

  function refresh(): void {
    const spec = readSpec();
    const runs = resultRuns(spec);
    // renderResult reuses `prev` as its surface host — don't clear it between calls.
    if (opts.renderResult) opts.renderResult(prev, runs);
    else prev.textContent = runs.map((r) => r.text).join("") || "(empty)";
    note.textContent =
      spec.type === "PAGE" ? "Live — resolves to the current page per page (i, I, a, A per format)."
      : spec.type === "NUMPAGES" ? "Live — resolves to the document's total page count."
      : spec.type === "DATE" || spec.type === "TIME" ? "Materialized now; refresh with Update Field."
      : spec.type === "REF" ? "Shows the bookmarked text; refresh with Update Field."
      : spec.type === "PAGEREF" ? "Shows the bookmark's page number; refresh with Update Field."
      : "Materialized from the comparison; refresh with Update Field.";
    // A cross-reference with no target bookmark can't be inserted.
    const noTarget = (spec.type === "REF" || spec.type === "PAGEREF") && spec.bookmark === "";
    apply.disabled = noTarget;
  }

  typeSel.addEventListener("change", rebuildForm);
  typeSel.disabled = editing; // editing keeps the field's type

  const spacer = el("div", "spacer");
  const cancel = el("button", "cw-fc-btn", "Cancel");
  const apply = el("button", "cw-fc-btn primary", editing ? "Apply" : "Insert");
  foot.append(spacer, cancel, apply);

  body.append(left, right);

  const handle: FieldConstructorHandle = { close: shell.close };
  cancel.addEventListener("click", shell.close);
  apply.addEventListener("click", () => { opts.onApply(readSpec()); shell.close(); });

  rebuildForm();
  return handle;
}
