// Size & Position — a small draggable, non-blocking dialog for numeric authoring of
// a selected drawing shape: exact width/height, rotation (degrees), and — for a
// floating (anchored) shape — the X/Y offset from its anchor origin. It fills the
// precision gap left by the drag-only handles: repeatable, exact layouts and angles
// you can type instead of eyeballing (UX finding B2).
//
// Pattern mirrors pageLayout.ts: createDialogShell + an in/cm unit toggle, numeric
// px-backed fields, Escape/×/Cancel to close, and one committed op on Apply
// (a single setShapeProps carrying size + rotation + anchor offset = one undo step).

import type { Command } from "../editor/state";
import type { ShapeBlock, ShapePropsPatch } from "@cw/shared";
import { setShapeProps } from "../editor/commands";
import { normalizeDeg } from "../input/objectController";
import { createDialogShell } from "./dialogShell";
import { formatUnit, unitToPx, type LengthUnit } from "./units";

/** The slice of the editor the dialog needs (the full Editor satisfies it). */
export interface ShapeSizePositionEditor {
  /** The currently object-selected drawing shape, or null. Read once on open. */
  getSelectedShape(): ShapeBlock | null;
  dispatch(cmd: Command): void;
  focus(): void;
}

export interface ShapeSizePositionOptions {
  editor: ShapeSizePositionEditor;
  onClose?: () => void;
}

export interface ShapeSizePositionHandle {
  close(): void;
}

const CSS = `
.cw-sp-backdrop{position:fixed;inset:0;z-index:1100;background:rgba(20,22,26,.30);}
.cw-sp-modal{position:fixed;width:min(360px,95vw);max-height:90vh;display:flex;flex-direction:column;background:#fff;border-radius:10px;
  box-shadow:0 18px 56px rgba(0,0,0,.34);font:13px/1.5 Arial,sans-serif;color:#202124;overflow:hidden;}
.cw-sp-head{display:flex;align-items:center;gap:10px;padding:12px 16px;border-bottom:1px solid #e6e8eb;cursor:move;}
.cw-sp-head h2{margin:0;font-size:15px;font-weight:600;flex:1 1 auto;}
.cw-sp-unit{display:flex;gap:4px;align-items:center;font-size:11px;color:#80868b;}
.cw-sp-unit select{height:24px;border:1px solid #d0d4d9;border-radius:5px;font:12px Arial;}
.cw-sp-x{border:none;background:transparent;font-size:20px;line-height:1;color:#5f6368;cursor:pointer;width:28px;height:28px;border-radius:6px;}
.cw-sp-x:hover{background:#e8eaed;}
.cw-sp-body{padding:14px 16px;display:flex;flex-direction:column;gap:12px;overflow:auto;}
.cw-sp-sect{font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:#80868b;margin:2px 0 -4px;}
.cw-sp-field{display:flex;flex-direction:column;gap:4px;}
.cw-sp-field label{font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:#80868b;}
.cw-sp-field input{height:30px;border:1px solid #d0d4d9;border-radius:6px;padding:0 8px;font:13px Arial,sans-serif;box-sizing:border-box;}
.cw-sp-field input:disabled{background:#f1f3f4;color:#9aa0a6;}
.cw-sp-row{display:flex;gap:8px;}.cw-sp-row>*{flex:1 1 0;min-width:0;}
.cw-sp-note{font-size:11px;color:#80868b;margin:-2px 0 0;}
.cw-sp-foot{display:flex;align-items:center;gap:8px;padding:11px 16px;border-top:1px solid #e6e8eb;}
.cw-sp-foot .spacer{flex:1 1 auto;}
.cw-sp-btn{height:30px;padding:0 14px;border:1px solid #d0d4d9;border-radius:6px;background:#fff;cursor:pointer;font-size:13px;color:#3c4043;}
.cw-sp-btn:hover{background:#f1f3f4;}
.cw-sp-btn.primary{border-color:#1a73e8;background:#1a73e8;color:#fff;}
.cw-sp-btn.primary:hover{background:#1864cc;}`;

const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string, text?: string): HTMLElementTagNameMap[K] => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
};

const labelled = (label: string, control: HTMLElement): HTMLElement => {
  const wrap = el("div", "cw-sp-field");
  wrap.append(el("label", undefined, label), control);
  return wrap;
};

const rowOf = (...kids: HTMLElement[]): HTMLElement => {
  const r = el("div", "cw-sp-row");
  r.append(...kids);
  return r;
};

const option = (value: string, label: string): HTMLOptionElement => {
  const o = el("option");
  o.value = value;
  o.textContent = label;
  return o;
};

/** A numeric input bound to a px value, shown in the active length unit. */
interface NumField {
  input: HTMLInputElement;
  getPx(): number;
  refreshUnit(): void;
}

export function showShapeSizePosition(opts: ShapeSizePositionOptions): ShapeSizePositionHandle {
  const { editor } = opts;
  const shape = editor.getSelectedShape();
  // Guard: opened without a shape selected — nothing to author. (The ribbon button
  // and menu entry only fire on a selected shape, so this is defensive.)
  if (!shape) {
    opts.onClose?.();
    return { close: () => {} };
  }
  const shapeId = shape.id;
  const anchored = !!shape.anchor;

  let unit: LengthUnit = "in";
  const numFields: NumField[] = [];
  const makeNum = (px0: number, min = 0, disabled = false): NumField => {
    let px = px0;
    const input = el("input");
    input.type = "number";
    input.step = "0.01";
    input.value = formatUnit(px, unit);
    input.disabled = disabled;
    input.addEventListener("input", () => {
      const v = Number(input.value);
      if (Number.isFinite(v)) px = Math.max(min, unitToPx(v, unit));
    });
    const nf: NumField = {
      input,
      getPx: () => px,
      refreshUnit: () => {
        input.value = formatUnit(px, unit);
      },
    };
    numFields.push(nf);
    return nf;
  };

  // ---- unit selector in the header -----------------------------------------
  const unitWrap = el("div", "cw-sp-unit");
  const unitSel = el("select");
  unitSel.append(option("in", "inches"), option("cm", "cm"));
  unitWrap.append(el("span", undefined, "Units"), unitSel);

  const shell = createDialogShell({
    prefix: "cw-sp",
    cssId: "cw-sp-styles",
    css: CSS,
    title: "Size & position",
    headExtras: [unitWrap],
    extraNoDrag: ".cw-sp-unit",
    ...(opts.onClose ? { onClose: opts.onClose } : {}),
  });

  // ---- fields ---------------------------------------------------------------
  const widthNum = makeNum(shape.widthPx, 1);
  const heightNum = makeNum(shape.heightPx, 1);

  const rotInput = el("input");
  rotInput.type = "number";
  rotInput.step = "1";
  rotInput.value = String(Math.round(normalizeDeg(shape.rotation ?? 0)));

  const xNum = makeNum(shape.anchor?.offsetXPx ?? 0, Number.NEGATIVE_INFINITY, !anchored);
  const yNum = makeNum(shape.anchor?.offsetYPx ?? 0, Number.NEGATIVE_INFINITY, !anchored);

  shell.body.append(
    el("div", "cw-sp-sect", "Size"),
    rowOf(labelled("Width", widthNum.input), labelled("Height", heightNum.input)),
    el("div", "cw-sp-sect", "Rotation"),
    labelled("Angle (degrees)", rotInput),
    el("div", "cw-sp-sect", "Position"),
    rowOf(labelled("Horizontal (X)", xNum.input), labelled("Vertical (Y)", yNum.input)),
  );
  if (!anchored) {
    shell.body.append(
      el("div", "cw-sp-note", "Position applies to floating shapes — set Wrap to “In Front of Text” or “Behind Text” to move by offset."),
    );
  }

  // ---- unit toggle ----------------------------------------------------------
  unitSel.addEventListener("change", () => {
    unit = unitSel.value as LengthUnit;
    for (const nf of numFields) nf.refreshUnit();
  });

  // ---- footer ---------------------------------------------------------------
  const spacer = el("div", "spacer");
  const cancel = el("button", "cw-sp-btn", "Cancel");
  const apply = el("button", "cw-sp-btn primary", "Apply");
  shell.foot.append(spacer, cancel, apply);
  cancel.addEventListener("click", shell.close);

  apply.addEventListener("click", () => {
    // Re-resolve the shape (selection could have changed under a still-open panel).
    const shp = editor.getSelectedShape();
    if (!shp || shp.id !== shapeId) {
      shell.close();
      return;
    }
    const rot = normalizeDeg(Number(rotInput.value) || 0);
    const patch: ShapePropsPatch = {
      widthPx: Math.max(1, Math.round(widthNum.getPx())),
      heightPx: Math.max(1, Math.round(heightNum.getPx())),
      // 0° stores as "no rotation" (null) to keep the model clean, matching the
      // rotate handle's commit.
      rotation: rot === 0 ? null : rot,
    };
    // Only floating shapes carry an anchor offset — fold X/Y into the same op so the
    // whole edit is a single undo step (mirrors moveAnchoredShape's anchor patch).
    if (shp.anchor) {
      patch.anchor = { ...shp.anchor, offsetXPx: Math.round(xNum.getPx()), offsetYPx: Math.round(yNum.getPx()) };
    }
    editor.dispatch(setShapeProps(shapeId, patch));
    shell.close();
    editor.focus();
  });

  return { close: shell.close };
}
