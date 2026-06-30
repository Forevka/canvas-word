// Paragraph dialog — Word's "Paragraph" / "Borders and Shading" for the caret
// paragraph(s). Unlike the table Borders & Shading panel (which applies each click
// live), this dialog gathers every control and commits ONE patch on OK, so the
// whole change is a single undo step. The border spec UI (color / width / style +
// per-edge toggles) mirrors tableProperties.ts.

import type { CellBorder, ParaBorders, ParaStyle } from "@cw/shared";
import { injectCssOnce } from "./styles";
import { makeFloatingDialog } from "./floatingDialog";

export type ParaBorderStyle = NonNullable<CellBorder["style"]>;
export type LineSpacingRule = "auto" | "exact" | "atLeast";
export type VerticalTextAlign = NonNullable<ParaStyle["textAlignment"]>;

/** Seed values for the dialog — the caret paragraph's current (direct) style. */
export interface ParagraphDialogInit {
  borders: ParaBorders | undefined;
  shading: string | null;
  contextualSpacing: boolean;
  lineRule: LineSpacingRule;
  /** Multiplier used when the rule is "auto". */
  lineHeight: number;
  /** Fixed line height in px, used when the rule is "exact"/"atLeast". */
  lineHeightPx: number;
  widowControl: boolean;
  textAlignment: VerticalTextAlign;
  mirrorIndents: boolean;
  suppressLineNumbers: boolean;
  adjustRightInd: boolean;
}

export interface ParagraphDialogCallbacks {
  /** Commit the gathered settings as one undoable paragraph-style patch. */
  apply(patch: Partial<ParaStyle>): void;
}

export interface ParagraphDialogHandle {
  close(): void;
}

const EDGES = ["top", "right", "bottom", "left", "between"] as const;
type Edge = (typeof EDGES)[number];
const EDGE_LABEL: Record<Edge, string> = { top: "Top", right: "Right", bottom: "Bottom", left: "Left", between: "Between" };
const STYLES: ParaBorderStyle[] = ["single", "double", "dashed", "dotted"];

const PDLG_CSS = `
.cw-pdlg-backdrop{position:fixed;inset:0;z-index:1100;pointer-events:none;}
.cw-pdlg-modal{position:fixed;width:min(440px,94vw);max-height:88vh;display:flex;flex-direction:column;
  background:#fff;border-radius:10px;box-shadow:0 18px 56px rgba(0,0,0,.34);border:1px solid #d9dce1;
  font:13px/1.5 Arial,sans-serif;color:#202124;overflow:hidden;pointer-events:auto;}
.cw-pdlg-head{display:flex;align-items:center;gap:10px;padding:11px 14px;border-bottom:1px solid #e6e8eb;background:#f7f8fa;}
.cw-pdlg-head h2{margin:0;font-size:14px;font-weight:600;flex:1 1 auto;}
.cw-pdlg-x{border:none;background:transparent;font-size:20px;line-height:1;color:#5f6368;cursor:pointer;
  width:28px;height:28px;border-radius:6px;}
.cw-pdlg-x:hover{background:#e8eaed;}
.cw-pdlg-body{padding:14px 16px;overflow:auto;display:flex;flex-direction:column;gap:16px;}
.cw-pdlg-section-title{font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:#80868b;margin:0 0 8px;}
.cw-pdlg-row{display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin-bottom:8px;}
.cw-pdlg-row label{display:flex;align-items:center;gap:6px;color:#5f6368;}
.cw-pdlg-row input[type=number]{width:64px;height:28px;border:1px solid #c8c6c4;border-radius:4px;padding:0 6px;}
.cw-pdlg-row select{height:28px;border:1px solid #c8c6c4;border-radius:4px;padding:0 6px;background:#fff;}
.cw-pdlg-swatch{width:28px;height:28px;border:1px solid #c8c6c4;border-radius:4px;padding:0;cursor:pointer;background:#000;}
.cw-pdlg-edges{display:flex;gap:14px;flex-wrap:wrap;}
.cw-pdlg-edges label{color:#3c4043;}
.cw-pdlg-preview{width:120px;height:60px;border:1px dashed #c8c6c4;border-radius:6px;
  display:grid;place-items:center;background:#fff;}
.cw-pdlg-preview .box{width:78px;height:34px;}
.cw-pdlg-foot{display:flex;justify-content:flex-end;gap:8px;padding:11px 16px;border-top:1px solid #e6e8eb;}
.cw-pdlg-btn{height:30px;padding:0 14px;border:1px solid #d0d4d9;border-radius:6px;background:#fff;cursor:pointer;
  font-size:13px;color:#3c4043;}
.cw-pdlg-btn:hover{background:#f1f3f4;}
.cw-pdlg-btn.primary{border-color:#1a73e8;background:#1a73e8;color:#fff;}
.cw-pdlg-btn.primary:hover{background:#1864cc;}`;

const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string): HTMLElementTagNameMap[K] => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  return e;
};

const cssBorderStyle = (s: ParaBorderStyle): string =>
  s === "double" ? "double" : s === "dashed" ? "dashed" : s === "dotted" ? "dotted" : "solid";

/** Native <input type=color> needs #rrggbb; coerce loose CSS colors. */
function toHexColor(c: string): string {
  if (/^#[0-9a-fA-F]{6}$/.test(c)) return c.toLowerCase();
  if (/^#[0-9a-fA-F]{3}$/.test(c)) {
    const r = c[1]!, g = c[2]!, b = c[3]!;
    return `#${r}${r}${g}${g}${b}${b}`.toLowerCase();
  }
  return "#000000";
}

/** A checkbox + label row helper. */
function checkRow(labelText: string, checked: boolean): { row: HTMLLabelElement; input: HTMLInputElement } {
  const row = el("label");
  const input = el("input");
  input.type = "checkbox";
  input.checked = checked;
  row.append(input, document.createTextNode(labelText));
  return { row, input };
}

export function showParagraphDialog(init: ParagraphDialogInit, cb: ParagraphDialogCallbacks): ParagraphDialogHandle {
  injectCssOnce("cw-pdlg-styles", PDLG_CSS);

  const backdrop = el("div", "cw-pdlg-backdrop");
  const modal = el("div", "cw-pdlg-modal");
  modal.addEventListener("mousedown", (e) => e.stopPropagation());

  // Header
  const head = el("div", "cw-pdlg-head");
  const h2 = el("h2");
  h2.textContent = "Paragraph";
  const xBtn = el("button", "cw-pdlg-x");
  xBtn.textContent = "×";
  xBtn.title = "Close (Esc)";
  head.append(h2, xBtn);

  const body = el("div", "cw-pdlg-body");

  // ---- Borders ------------------------------------------------------------
  // Seed the spec from whichever edge already exists, else a 1px black single line.
  const seedEdge = init.borders && (init.borders.top ?? init.borders.right ?? init.borders.bottom ?? init.borders.left ?? init.borders.between);
  let color = seedEdge?.color ?? "#000000";
  let widthPx = seedEdge?.widthPx ?? 1;
  let styleName: ParaBorderStyle = seedEdge?.style ?? "single";

  const bSection = el("div");
  const bTitle = el("div", "cw-pdlg-section-title");
  bTitle.textContent = "Borders";
  bSection.appendChild(bTitle);

  const specRow = el("div", "cw-pdlg-row");
  const colorInput = el("input");
  colorInput.type = "color";
  colorInput.className = "cw-pdlg-swatch";
  colorInput.value = toHexColor(color);
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

  const preview = el("div", "cw-pdlg-preview");
  const previewBox = el("div", "box");
  preview.appendChild(previewBox);
  const refreshPreview = (): void => {
    previewBox.style.border = `${Math.max(1, widthPx)}px ${cssBorderStyle(styleName)} ${color}`;
  };
  colorInput.addEventListener("input", () => { color = colorInput.value; refreshPreview(); });
  widthInput.addEventListener("change", () => {
    const v = Number(widthInput.value);
    if (Number.isFinite(v) && v > 0) widthPx = v;
    widthInput.value = String(widthPx);
    refreshPreview();
  });
  styleSelect.addEventListener("change", () => { styleName = (styleSelect.value as ParaBorderStyle) || "single"; refreshPreview(); });
  specRow.append(colorLabel, widthLabel, styleLabel, preview);
  refreshPreview();

  // Per-edge toggles (which edges of the box are drawn).
  const edgeRow = el("div", "cw-pdlg-edges");
  const edgeInputs: Record<Edge, HTMLInputElement> = {} as Record<Edge, HTMLInputElement>;
  for (const e of EDGES) {
    const { row, input } = checkRow(EDGE_LABEL[e], !!init.borders?.[e]);
    edgeInputs[e] = input;
    edgeRow.append(row);
  }
  bSection.append(specRow, edgeRow);

  // ---- Shading ------------------------------------------------------------
  const sSection = el("div");
  const sTitle = el("div", "cw-pdlg-section-title");
  sTitle.textContent = "Shading (fill)";
  const sRow = el("div", "cw-pdlg-row");
  const { row: fillToggleRow, input: fillToggle } = checkRow("Fill", init.shading !== null);
  const fillInput = el("input");
  fillInput.type = "color";
  fillInput.className = "cw-pdlg-swatch";
  fillInput.value = toHexColor(init.shading ?? "#ffff00");
  sRow.append(fillToggleRow, fillInput);
  sSection.append(sTitle, sRow);

  // ---- Spacing + flags ----------------------------------------------------
  const oSection = el("div");
  const oTitle = el("div", "cw-pdlg-section-title");
  oTitle.textContent = "Spacing & line breaks";

  const lsRow = el("div", "cw-pdlg-row");
  const ruleSelect = el("select");
  for (const [val, lbl] of [["auto", "Multiple"], ["atLeast", "At least"], ["exact", "Exactly"]] as const) {
    const o = el("option");
    o.value = val;
    o.textContent = lbl;
    if (val === init.lineRule) o.selected = true;
    ruleSelect.appendChild(o);
  }
  const ruleLabel = el("label");
  ruleLabel.append("Line spacing", ruleSelect);
  // One numeric field, reinterpreted by the rule: a multiplier for "auto", px for fixed.
  const multInput = el("input");
  multInput.type = "number";
  multInput.min = "0.5";
  multInput.step = "0.05";
  const pxInput = el("input");
  pxInput.type = "number";
  pxInput.min = "1";
  pxInput.step = "1";
  multInput.value = String(init.lineHeight || 1);
  pxInput.value = String(Math.round(init.lineHeightPx || 16));
  const multLabel = el("label");
  multLabel.append("At", multInput);
  const pxLabel = el("label");
  pxLabel.append("At (px)", pxInput);
  const syncRuleInputs = (): void => {
    const auto = ruleSelect.value === "auto";
    multLabel.style.display = auto ? "" : "none";
    pxLabel.style.display = auto ? "none" : "";
  };
  ruleSelect.addEventListener("change", syncRuleInputs);
  lsRow.append(ruleLabel, multLabel, pxLabel);
  syncRuleInputs();

  const vaRow = el("div", "cw-pdlg-row");
  const vaSelect = el("select");
  for (const [val, lbl] of [["baseline", "Baseline"], ["top", "Top"], ["center", "Center"], ["bottom", "Bottom"]] as const) {
    const o = el("option");
    o.value = val;
    o.textContent = lbl;
    if (val === init.textAlignment) o.selected = true;
    vaSelect.appendChild(o);
  }
  const vaLabel = el("label");
  vaLabel.append("Vertical alignment", vaSelect);
  vaRow.append(vaLabel);

  const flagRow = el("div", "cw-pdlg-edges");
  const { row: csRow, input: csInput } = checkRow("Contextual spacing", init.contextualSpacing);
  const { row: wcRow, input: wcInput } = checkRow("Widow/orphan control", init.widowControl);
  const { row: miRow, input: miInput } = checkRow("Mirror indents", init.mirrorIndents);
  const { row: slRow, input: slInput } = checkRow("Suppress line numbers", init.suppressLineNumbers);
  const { row: arRow, input: arInput } = checkRow("Adjust right indent", init.adjustRightInd);
  flagRow.append(csRow, wcRow, miRow, slRow, arRow);

  oSection.append(oTitle, lsRow, vaRow, flagRow);

  body.append(bSection, sSection, oSection);

  // Footer
  const foot = el("div", "cw-pdlg-foot");
  const cancelBtn = el("button", "cw-pdlg-btn");
  cancelBtn.textContent = "Cancel";
  const okBtn = el("button", "cw-pdlg-btn primary");
  okBtn.textContent = "OK";
  foot.append(cancelBtn, okBtn);

  modal.append(head, body, foot);
  backdrop.appendChild(modal);
  document.body.appendChild(backdrop);

  const ac = new AbortController();
  const handle: ParagraphDialogHandle = {
    close(): void {
      backdrop.remove();
      ac.abort();
    },
  };
  makeFloatingDialog({ backdrop, modal, handle: head, signal: ac.signal, noDrag: ".cw-pdlg-x" });

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
  cancelBtn.addEventListener("click", () => handle.close());

  okBtn.addEventListener("click", () => {
    const edge = (): CellBorder => ({ color, widthPx, ...(styleName !== "single" ? { style: styleName } : {}) });
    const borders: ParaBorders = {};
    for (const e of EDGES) if (edgeInputs[e].checked) borders[e] = edge();
    const hasBorder = Object.keys(borders).length > 0;

    // The dialog represents the FULL desired paragraph state, so each control's value
    // is written (an unchecked flag stores the default; `baseline` is the textAlignment
    // default). `borders`/`shading` accept `undefined` (they're declared `| undefined`)
    // to clear; the boolean flags carry their literal value.
    const patch: Partial<ParaStyle> = {
      borders: hasBorder ? borders : undefined,
      shading: fillToggle.checked ? fillInput.value : undefined,
      contextualSpacing: csInput.checked,
      widowControl: wcInput.checked,
      mirrorIndents: miInput.checked,
      suppressLineNumbers: slInput.checked,
      adjustRightInd: arInput.checked,
      textAlignment: vaSelect.value as VerticalTextAlign,
    };

    const rule = ruleSelect.value as LineSpacingRule;
    if (rule === "auto") {
      const m = Number(multInput.value);
      if (Number.isFinite(m) && m > 0) patch.lineHeight = m;
      // Clear any fixed rule. `lineRule`/`lineHeightPx` aren't declared `| undefined`,
      // so under exactOptionalPropertyTypes they can't be assigned `undefined` via a
      // literal; setParaStyle spreads the patch, so an explicit-undefined key removes
      // the field. Object.assign sidesteps the literal-assignability check.
      Object.assign(patch, { lineRule: undefined, lineHeightPx: undefined });
    } else {
      const px = Number(pxInput.value);
      patch.lineRule = rule;
      if (Number.isFinite(px) && px > 0) patch.lineHeightPx = px;
    }

    cb.apply(patch);
    handle.close();
  });

  return handle;
}
