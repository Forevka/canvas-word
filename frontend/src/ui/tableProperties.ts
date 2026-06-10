// Table "Borders & Shading" modal — Word's dialog of the same name, scoped to the
// current cell selection (or the caret's cell). Pure presentation: index.ts
// captures the target range and supplies the apply callbacks; edits are applied
// live (each click is its own undo step), so there's no separate OK/commit step.

import type { CellBorder } from "@cw/shared";
import type { BorderEdgeFlags } from "../editor/commands";

export type BorderStyleName = NonNullable<CellBorder["style"]> | "single";

export interface TablePropertiesInit {
  /** Seed for the border controls (from the top-left cell of the range). */
  color: string;
  widthPx: number;
  style: BorderStyleName;
  /** Current shading of the top-left cell, or null for no fill. */
  shading: string | null;
  /** Human label for the target, e.g. "4 cells" / "1 cell". */
  rangeLabel: string;
  /** Interior presets (Inside / Inside H / Inside V) only apply to a multi-cell range. */
  multiCell: boolean;
}

export interface TablePropertiesCallbacks {
  applyBorders(spec: CellBorder | null, edges: BorderEdgeFlags): void;
  applyShading(fill: string | null): void;
}

export interface TablePropertiesHandle {
  close(): void;
}

const STYLES: BorderStyleName[] = ["single", "double", "dashed", "dotted"];

let cssInjected = false;
function injectCss(): void {
  if (cssInjected) return;
  cssInjected = true;
  const style = document.createElement("style");
  style.textContent = `
.cw-tbl-backdrop{position:fixed;inset:0;z-index:1100;background:rgba(20,22,26,.38);
  display:flex;align-items:center;justify-content:center;}
.cw-tbl-modal{width:min(560px,94vw);max-height:90vh;display:flex;flex-direction:column;
  background:#fff;border-radius:10px;box-shadow:0 18px 56px rgba(0,0,0,.34);
  font:13px/1.5 Arial,sans-serif;color:#202124;overflow:hidden;}
.cw-tbl-head{display:flex;align-items:center;gap:10px;padding:13px 16px;border-bottom:1px solid #e6e8eb;}
.cw-tbl-head h2{margin:0;font-size:15px;font-weight:600;flex:1 1 auto;}
.cw-tbl-badge{font-size:11px;font-weight:600;color:#0b57d0;background:#e8f0fe;border-radius:10px;padding:2px 9px;}
.cw-tbl-x{border:none;background:transparent;font-size:20px;line-height:1;color:#5f6368;cursor:pointer;
  width:28px;height:28px;border-radius:6px;}
.cw-tbl-x:hover{background:#e8eaed;}
.cw-tbl-body{padding:14px 16px;overflow:auto;display:flex;flex-direction:column;gap:16px;}
.cw-tbl-section-title{font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:#80868b;margin:0 0 8px;}
.cw-tbl-row{display:flex;align-items:center;gap:10px;flex-wrap:wrap;}
.cw-tbl-row label{display:flex;align-items:center;gap:6px;color:#5f6368;}
.cw-tbl-spec{display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin-bottom:10px;}
.cw-tbl-spec input[type=number]{width:56px;height:28px;border:1px solid #c8c6c4;border-radius:4px;padding:0 6px;}
.cw-tbl-spec select{height:28px;border:1px solid #c8c6c4;border-radius:4px;padding:0 6px;background:#fff;}
.cw-tbl-swatch{width:28px;height:28px;border:1px solid #c8c6c4;border-radius:4px;padding:0;cursor:pointer;background:#000;}
.cw-tbl-btn{height:30px;padding:0 12px;border:1px solid #d0d4d9;border-radius:6px;background:#fff;cursor:pointer;
  font-size:13px;color:#3c4043;display:inline-flex;align-items:center;gap:6px;}
.cw-tbl-btn:hover{background:#f1f3f4;}
.cw-tbl-btn:disabled{opacity:.45;cursor:default;}
.cw-tbl-preview{width:108px;height:72px;border:1px dashed #c8c6c4;border-radius:6px;align-self:center;
  display:grid;place-items:center;background:#fff;}
.cw-tbl-preview .box{width:64px;height:40px;}
.cw-tbl-foot{display:flex;justify-content:flex-end;gap:8px;padding:11px 16px;border-top:1px solid #e6e8eb;}
.cw-tbl-foot .cw-tbl-btn.primary{border-color:#1a73e8;background:#1a73e8;color:#fff;}
.cw-tbl-foot .cw-tbl-btn.primary:hover{background:#1864cc;}`;
  document.head.appendChild(style);
}

const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string): HTMLElementTagNameMap[K] => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  return e;
};

/** CSS border-style for a model border style name (double/dashed/dotted/solid). */
const cssBorderStyle = (s: BorderStyleName): string =>
  s === "double" ? "double" : s === "dashed" ? "dashed" : s === "dotted" ? "dotted" : "solid";

export function showTableProperties(init: TablePropertiesInit, cb: TablePropertiesCallbacks): TablePropertiesHandle {
  injectCss();

  // Live border spec the buttons apply.
  let color = init.color;
  let widthPx = init.widthPx;
  let styleName: BorderStyleName = init.style;
  const spec = (): CellBorder => ({ color, widthPx, ...(styleName !== "single" ? { style: styleName } : {}) });

  const backdrop = el("div", "cw-tbl-backdrop");
  const modal = el("div", "cw-tbl-modal");
  modal.addEventListener("mousedown", (e) => e.stopPropagation());

  // Header
  const head = el("div", "cw-tbl-head");
  const h2 = el("h2");
  h2.textContent = "Borders & Shading";
  const badge = el("span", "cw-tbl-badge");
  badge.textContent = init.rangeLabel;
  const xBtn = el("button", "cw-tbl-x");
  xBtn.textContent = "×";
  xBtn.title = "Close (Esc)";
  head.append(h2, badge, xBtn);

  const body = el("div", "cw-tbl-body");

  // ---- Borders section ----------------------------------------------------
  const bSection = el("div");
  const bTitle = el("div", "cw-tbl-section-title");
  bTitle.textContent = "Borders";
  bSection.appendChild(bTitle);

  // Spec controls: color, width, style + a live preview.
  const specRow = el("div", "cw-tbl-spec");
  const colorInput = el("input");
  colorInput.type = "color";
  colorInput.className = "cw-tbl-swatch";
  colorInput.value = toHexColor(color);
  colorInput.title = "Line color";
  const colorLabel = el("label");
  colorLabel.append("Color", colorInput);

  const widthInput = el("input");
  widthInput.type = "number";
  widthInput.min = "0.25";
  widthInput.max = "12";
  widthInput.step = "0.25";
  widthInput.value = String(widthPx);
  const widthLabel = el("label");
  widthLabel.append("Width (px)", widthInput);

  const styleSelect = el("select");
  for (const s of STYLES) {
    const opt = el("option");
    opt.value = s;
    opt.textContent = s[0]!.toUpperCase() + s.slice(1);
    if (s === styleName) opt.selected = true;
    styleSelect.appendChild(opt);
  }
  const styleLabel = el("label");
  styleLabel.append("Style", styleSelect);

  const preview = el("div", "cw-tbl-preview");
  const previewBox = el("div", "box");
  preview.appendChild(previewBox);
  const refreshPreview = (): void => {
    previewBox.style.border = `${Math.max(1, widthPx)}px ${cssBorderStyle(styleName)} ${color}`;
  };

  colorInput.addEventListener("input", () => {
    color = colorInput.value;
    refreshPreview();
  });
  widthInput.addEventListener("change", () => {
    const v = Number(widthInput.value);
    if (Number.isFinite(v) && v > 0) widthPx = v;
    widthInput.value = String(widthPx);
    refreshPreview();
  });
  styleSelect.addEventListener("change", () => {
    styleName = (styleSelect.value as BorderStyleName) || "single";
    refreshPreview();
  });
  specRow.append(colorLabel, widthLabel, styleLabel, preview);
  refreshPreview();

  // Presets — apply immediately with the current spec.
  const presetRow = el("div", "cw-tbl-row");
  const allFlags: BorderEdgeFlags = { top: true, right: true, bottom: true, left: true, insideH: true, insideV: true };
  const presetBtn = (label: string, run: () => void, enabled = true): HTMLButtonElement => {
    const b = el("button", "cw-tbl-btn");
    b.textContent = label;
    b.disabled = !enabled;
    b.addEventListener("click", run);
    return b;
  };
  presetRow.append(
    presetBtn("All", () => cb.applyBorders(spec(), allFlags)),
    presetBtn("Outside", () => cb.applyBorders(spec(), { top: true, right: true, bottom: true, left: true })),
    presetBtn("Inside", () => cb.applyBorders(spec(), { insideH: true, insideV: true }), init.multiCell),
    presetBtn("None", () => cb.applyBorders(null, allFlags)),
  );

  // Individual edges.
  const edgeRow = el("div", "cw-tbl-row");
  const edgeBtn = (label: string, flag: BorderEdgeFlags, enabled = true): HTMLButtonElement =>
    presetBtn(label, () => cb.applyBorders(spec(), flag), enabled);
  edgeRow.append(
    edgeBtn("Top", { top: true }),
    edgeBtn("Bottom", { bottom: true }),
    edgeBtn("Left", { left: true }),
    edgeBtn("Right", { right: true }),
    edgeBtn("Inside H", { insideH: true }, init.multiCell),
    edgeBtn("Inside V", { insideV: true }, init.multiCell),
  );

  bSection.append(specRow, presetRow, edgeRow);

  // ---- Shading section ----------------------------------------------------
  const sSection = el("div");
  const sTitle = el("div", "cw-tbl-section-title");
  sTitle.textContent = "Shading (fill)";
  const sRow = el("div", "cw-tbl-row");
  const fillInput = el("input");
  fillInput.type = "color";
  fillInput.className = "cw-tbl-swatch";
  fillInput.value = toHexColor(init.shading ?? "#ffffff");
  const fillLabel = el("label");
  fillLabel.append("Fill", fillInput);
  const applyFill = presetBtn("Apply Fill", () => cb.applyShading(fillInput.value));
  const noFill = presetBtn("No Fill", () => cb.applyShading(null));
  sRow.append(fillLabel, applyFill, noFill);
  sSection.append(sTitle, sRow);

  body.append(bSection, sSection);

  // Footer
  const foot = el("div", "cw-tbl-foot");
  const doneBtn = el("button", "cw-tbl-btn primary");
  doneBtn.textContent = "Done";
  foot.append(doneBtn);

  modal.append(head, body, foot);
  backdrop.appendChild(modal);
  document.body.appendChild(backdrop);

  const handle: TablePropertiesHandle = {
    close(): void {
      backdrop.remove();
      window.removeEventListener("keydown", onKey, true);
    },
  };
  const onKey = (ev: KeyboardEvent): void => {
    if (ev.key === "Escape") {
      ev.preventDefault();
      ev.stopPropagation();
      handle.close();
    }
  };
  window.addEventListener("keydown", onKey, true);
  backdrop.addEventListener("mousedown", () => handle.close());
  xBtn.addEventListener("click", () => handle.close());
  doneBtn.addEventListener("click", () => handle.close());
  return handle;
}

/** Native <input type=color> needs a #rrggbb value; coerce loose CSS colors. */
function toHexColor(c: string): string {
  if (/^#[0-9a-fA-F]{6}$/.test(c)) return c.toLowerCase();
  if (/^#[0-9a-fA-F]{3}$/.test(c)) {
    const r = c[1]!, g = c[2]!, b = c[3]!;
    return `#${r}${r}${g}${g}${b}${b}`.toLowerCase();
  }
  return "#000000";
}
