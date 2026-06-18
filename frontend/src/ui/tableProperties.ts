// Table "Borders & Shading" modal — Word's dialog of the same name, scoped to the
// current cell selection (or the caret's cell). Pure presentation: index.ts
// captures the target range and supplies the apply callbacks; edits are applied
// live (each click is its own undo step), so there's no separate OK/commit step.

import type { CellBorder } from "@cw/shared";
import type { BorderEdgeFlags } from "../editor/commands";
import { injectCssOnce } from "./styles";

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

const TBL_CSS = `
/* Non-blocking layer: the panel floats but the document stays visible + interactive
   so border/shading edits are seen live; clicks outside don't close it. */
.cw-tbl-backdrop{position:fixed;inset:0;z-index:1100;pointer-events:none;}
.cw-tbl-modal{position:fixed;width:min(420px,94vw);max-height:88vh;display:flex;flex-direction:column;
  background:#fff;border-radius:10px;box-shadow:0 18px 56px rgba(0,0,0,.34);border:1px solid #d9dce1;
  font:13px/1.5 Arial,sans-serif;color:#202124;overflow:hidden;pointer-events:auto;}
.cw-tbl-head{display:flex;align-items:center;gap:10px;padding:11px 14px;border-bottom:1px solid #e6e8eb;
  cursor:move;user-select:none;background:#f7f8fa;}
.cw-tbl-head::before{content:"⠿";color:#b0b4ba;font-size:14px;line-height:1;margin-right:-2px;}
.cw-tbl-head h2{margin:0;font-size:14px;font-weight:600;flex:1 1 auto;}
.cw-tbl-hint{font-size:11px;color:#9aa0a6;margin:-4px 0 2px;}
.cw-tbl-caption{font-size:11px;color:#5f6368;font-weight:600;margin:0 0 4px;}
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

const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string): HTMLElementTagNameMap[K] => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  return e;
};

/** CSS border-style for a model border style name (double/dashed/dotted/solid). */
const cssBorderStyle = (s: BorderStyleName): string =>
  s === "double" ? "double" : s === "dashed" ? "dashed" : s === "dotted" ? "dotted" : "solid";

export function showTableProperties(init: TablePropertiesInit, cb: TablePropertiesCallbacks): TablePropertiesHandle {
  injectCssOnce("cw-tbl-styles", TBL_CSS);

  // Live border spec the buttons apply.
  let color = init.color;
  let widthPx = init.widthPx;
  let styleName: BorderStyleName = init.style;
  const spec = (): CellBorder => ({ color, widthPx, ...(styleName !== "single" ? { style: styleName } : {}) });

  // Apply happens live when the user clicks a preset/edge (or Apply Fill). But a
  // user who just tweaks the spec (e.g. Width=5) and clicks Done expects that to
  // take effect — so Done commits any spec/fill change they made without having
  // clicked a preset. These flags track which case we're in.
  let borderSpecTouched = false; // color/width/style changed
  let bordersApplied = false; // a preset/edge button was clicked
  let fillTouched = false; // fill color changed
  let shadingApplied = false; // Apply Fill / No Fill was clicked
  const doBorders = (s: CellBorder | null, edges: BorderEdgeFlags): void => {
    bordersApplied = true;
    cb.applyBorders(s, edges);
  };
  const doShading = (fill: string | null): void => {
    shadingApplied = true;
    cb.applyShading(fill);
  };

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
    borderSpecTouched = true;
    refreshPreview();
  });
  widthInput.addEventListener("change", () => {
    const v = Number(widthInput.value);
    if (Number.isFinite(v) && v > 0) widthPx = v;
    widthInput.value = String(widthPx);
    borderSpecTouched = true;
    refreshPreview();
  });
  styleSelect.addEventListener("change", () => {
    styleName = (styleSelect.value as BorderStyleName) || "single";
    borderSpecTouched = true;
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
    presetBtn("All", () => doBorders(spec(), allFlags)),
    presetBtn("Outside", () => doBorders(spec(), { top: true, right: true, bottom: true, left: true })),
    presetBtn("Inside", () => doBorders(spec(), { insideH: true, insideV: true }), init.multiCell),
    presetBtn("None", () => doBorders(null, allFlags)),
  );

  // Individual edges.
  const edgeRow = el("div", "cw-tbl-row");
  const edgeBtn = (label: string, flag: BorderEdgeFlags, enabled = true): HTMLButtonElement =>
    presetBtn(label, () => doBorders(spec(), flag), enabled);
  edgeRow.append(
    edgeBtn("Top", { top: true }),
    edgeBtn("Bottom", { bottom: true }),
    edgeBtn("Left", { left: true }),
    edgeBtn("Right", { right: true }),
    edgeBtn("Inside H", { insideH: true }, init.multiCell),
    edgeBtn("Inside V", { insideV: true }, init.multiCell),
  );

  const hint = el("div", "cw-tbl-hint");
  hint.textContent = "Pick a target to apply now — or set the options above and click Done.";
  const applyToCap = el("div", "cw-tbl-caption");
  applyToCap.textContent = "Apply borders to";
  const edgesCap = el("div", "cw-tbl-caption");
  edgesCap.textContent = "Individual edges";
  bSection.append(specRow, hint, applyToCap, presetRow, edgesCap, edgeRow);

  // ---- Shading section ----------------------------------------------------
  const sSection = el("div");
  const sTitle = el("div", "cw-tbl-section-title");
  sTitle.textContent = "Shading (fill)";
  const sRow = el("div", "cw-tbl-row");
  const fillInput = el("input");
  fillInput.type = "color";
  fillInput.className = "cw-tbl-swatch";
  fillInput.value = toHexColor(init.shading ?? "#ffffff");
  fillInput.addEventListener("input", () => { fillTouched = true; });
  const fillLabel = el("label");
  fillLabel.append("Fill", fillInput);
  const applyFill = presetBtn("Apply Fill", () => doShading(fillInput.value));
  const noFill = presetBtn("No Fill", () => doShading(null));
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

  // Float to the right edge by default so it doesn't cover the (often left-aligned)
  // table being edited; the user drags it wherever they like to watch live changes.
  const place = (left: number, top: number): void => {
    const maxLeft = Math.max(4, window.innerWidth - modal.offsetWidth - 6);
    const maxTop = Math.max(4, window.innerHeight - 44);
    modal.style.left = `${Math.min(Math.max(4, left), maxLeft)}px`;
    modal.style.top = `${Math.min(Math.max(4, top), maxTop)}px`;
  };
  place(window.innerWidth - modal.offsetWidth - 24, 72);

  const ac = new AbortController();
  const handle: TablePropertiesHandle = {
    close(): void {
      backdrop.remove();
      ac.abort(); // detaches every listener registered with ac.signal
    },
  };

  // Drag by the header (grab anywhere but the close button).
  let drag: { dx: number; dy: number } | null = null;
  head.addEventListener("mousedown", (ev) => {
    if ((ev.target as HTMLElement).closest(".cw-tbl-x")) return;
    const r = modal.getBoundingClientRect();
    drag = { dx: ev.clientX - r.left, dy: ev.clientY - r.top };
    ev.preventDefault();
  });
  window.addEventListener("mousemove", (ev) => { if (drag) place(ev.clientX - drag.dx, ev.clientY - drag.dy); }, { signal: ac.signal });
  window.addEventListener("mouseup", () => { drag = null; }, { signal: ac.signal });

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
  xBtn.addEventListener("click", () => handle.close());
  // Done commits spec/fill the user adjusted but didn't apply via a preset — so
  // "set Width=5 → Done" borders the selection (All) at 5px, as expected. If they
  // already clicked a preset/edge (live), Done just closes without re-applying.
  doneBtn.addEventListener("click", () => {
    if (borderSpecTouched && !bordersApplied) doBorders(spec(), allFlags);
    if (fillTouched && !shadingApplied) doShading(fillInput.value);
    handle.close();
  });
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
