// Text-style editor pane — the form half of the Style Manager for the two
// text-level style kinds: paragraph styles (character + paragraph properties)
// and character styles (character properties only). It is pure presentation:
// it renders controls into a host and exposes read() → NamedStyleSpec; the
// Style Manager turns that into an upsertNamedStyle command and owns the preview.
//
// "Defined vs inherited": a style stores only the fields it overrides (a Partial),
// the rest inherit through basedOn. Each control is SEEDED with the resolved value
// (so the user sees the real look) but only contributes to read() once it is in the
// style's own partial OR the user touches it — so opening and saving an unchanged
// style preserves its inheritance instead of flattening the whole cascade onto it.

import type { CharStyle, ParaStyle } from "@cw/shared";
import type { NamedStyleSpec } from "../editor/commands";
import { TOOLBAR_FONTS } from "../fonts/clones";
import { injectCssOnce } from "./styles";

export interface StyleEditorController {
  read(): NamedStyleSpec;
  onChange(cb: () => void): void;
  dispose(): void;
}

export interface TextStyleEditorInit {
  /** Identity + the style's OWN partials (not resolved). New → blank partials. */
  spec: NamedStyleSpec;
  /** Resolved (inherited + own) values, used to seed control display. */
  resolvedChar: Partial<CharStyle>;
  resolvedPara: Partial<ParaStyle>;
  /** Valid "based on" choices (same kind, minus self + descendants). */
  basedOnOptions: { id: string; name: string }[];
  isNew: boolean;
}

const CSS = `
.cw-se{display:flex;flex-direction:column;gap:14px;}
.cw-se-grp{display:flex;flex-direction:column;gap:9px;}
.cw-se-grp>h3{margin:0;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#80868b;font-weight:600;}
.cw-se-field{display:flex;flex-direction:column;gap:4px;}
.cw-se-field>label{font-size:11px;color:#5f6368;}
.cw-se-field input[type=text],.cw-se-field input[type=number],.cw-se-field select{
  height:30px;border:1px solid #d0d4d9;border-radius:6px;padding:0 8px;font:13px Arial,sans-serif;background:#fff;}
.cw-se-row{display:flex;gap:8px;flex-wrap:wrap;}
.cw-se-row>*{flex:1 1 0;min-width:90px;}
.cw-se-tgls{display:flex;gap:6px;}
.cw-se-tgl{width:32px;height:30px;border:1px solid #d0d4d9;border-radius:6px;background:#fff;cursor:pointer;font-size:14px;color:#3c4043;}
.cw-se-tgl.active{background:#1a73e8;border-color:#1a73e8;color:#fff;}
.cw-se-tgl[data-k=bold]{font-weight:700;}
.cw-se-tgl[data-k=italic]{font-style:italic;}
.cw-se-tgl[data-k=underline]{text-decoration:underline;}
.cw-se-tgl[data-k=strikethrough]{text-decoration:line-through;}
.cw-se-color{display:flex;align-items:center;gap:6px;}
.cw-se-color input[type=color]{width:34px;height:30px;border:1px solid #d0d4d9;border-radius:6px;padding:0;background:#fff;cursor:pointer;}
.cw-se-color button{height:26px;border:1px solid #d0d4d9;border-radius:6px;background:#fff;cursor:pointer;font-size:12px;color:#3c4043;padding:0 8px;}
.cw-se-checks{display:flex;flex-direction:column;gap:5px;}
.cw-se-checks label{display:flex;align-items:center;gap:7px;font-size:12px;color:#3c4043;}`;

const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string, text?: string): HTMLElementTagNameMap[K] => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
};
const field = (label: string, control: HTMLElement): HTMLElement => {
  const w = el("div", "cw-se-field");
  w.append(el("label", undefined, label), control);
  return w;
};
const option = (sel: HTMLSelectElement, value: string, label: string): void => {
  const o = el("option");
  o.value = value;
  o.textContent = label;
  sel.append(o);
};

const pxToPt = (px: number): number => Math.round(px * 0.75 * 2) / 2;
const ptToPx = (pt: number): number => (pt * 4) / 3;
const toHex = (c: string | undefined): string => {
  if (!c) return "#000000";
  if (/^#[0-9a-f]{6}$/i.test(c)) return c;
  const m = document.createElement("canvas").getContext("2d");
  if (m) {
    m.fillStyle = "#000";
    m.fillStyle = c;
    if (/^#[0-9a-f]{6}$/i.test(m.fillStyle)) return m.fillStyle;
  }
  return "#000000";
};

const ALIGNS: ParaStyle["align"][] = ["left", "center", "right", "justify"];

export function mountTextStyleEditor(host: HTMLElement, init: TextStyleEditorInit): StyleEditorController {
  injectCssOnce("cw-se-styles", CSS);
  const isPara = init.spec.type === "paragraph";
  const ownChar = init.spec.char ?? {};
  const ownPara = init.spec.para ?? {};
  const ch = { ...init.resolvedChar, ...ownChar };
  const pa = { ...init.resolvedPara, ...ownPara };
  const definedChar = new Set<keyof CharStyle>(Object.keys(ownChar) as (keyof CharStyle)[]);
  const definedPara = new Set<keyof ParaStyle>(Object.keys(ownPara) as (keyof ParaStyle)[]);

  let onChangeCb: () => void = () => {};
  const fire = (): void => onChangeCb();
  const markC = (k: keyof CharStyle): void => void definedChar.add(k);
  const markP = (k: keyof ParaStyle): void => void definedPara.add(k);

  const root = el("div", "cw-se");

  // -- Identity ---------------------------------------------------------------
  const idGrp = el("div", "cw-se-grp");
  const nameInput = el("input");
  nameInput.type = "text";
  nameInput.value = init.spec.name;
  nameInput.placeholder = "Style name";
  nameInput.addEventListener("input", fire);
  idGrp.append(field("Name", nameInput));

  const basedSel = el("select");
  option(basedSel, "", "(no base style)");
  for (const o of init.basedOnOptions) option(basedSel, o.id, o.name);
  basedSel.value = init.spec.basedOn ?? "";
  basedSel.addEventListener("change", fire);
  idGrp.append(field("Based on", basedSel));
  root.append(idGrp);

  // -- Character --------------------------------------------------------------
  const cGrp = el("div", "cw-se-grp");
  cGrp.append(el("h3", undefined, "Font"));

  const fontSel = el("select");
  for (const f of TOOLBAR_FONTS) option(fontSel, f.value, f.label);
  if (ch.fontFamily && !TOOLBAR_FONTS.some((f) => f.value === ch.fontFamily)) option(fontSel, ch.fontFamily, ch.fontFamily);
  fontSel.value = ch.fontFamily ?? TOOLBAR_FONTS[0]!.value;
  fontSel.addEventListener("change", () => { markC("fontFamily"); fire(); });

  const sizeInput = el("input");
  sizeInput.type = "number";
  sizeInput.min = "1";
  sizeInput.max = "96";
  sizeInput.step = "0.5";
  sizeInput.value = String(pxToPt(ch.fontSizePx ?? 16));
  sizeInput.addEventListener("input", () => { markC("fontSizePx"); fire(); });

  const fontRow = el("div", "cw-se-row");
  fontRow.append(field("Family", fontSel), field("Size (pt)", sizeInput));
  cGrp.append(fontRow);

  const tgls = el("div", "cw-se-tgls");
  const tglState: Record<"bold" | "italic" | "underline" | "strikethrough", boolean> = {
    bold: ch.bold ?? false, italic: ch.italic ?? false, underline: ch.underline ?? false, strikethrough: ch.strikethrough ?? false,
  };
  const glyph = { bold: "B", italic: "I", underline: "U", strikethrough: "S" } as const;
  for (const k of ["bold", "italic", "underline", "strikethrough"] as const) {
    const b = el("button", "cw-se-tgl", glyph[k]);
    b.type = "button";
    b.dataset.k = k;
    b.title = k;
    if (tglState[k]) b.classList.add("active");
    b.addEventListener("click", () => {
      tglState[k] = !tglState[k];
      b.classList.toggle("active", tglState[k]);
      markC(k);
      fire();
    });
    tgls.append(b);
  }
  cGrp.append(field("Style", tgls));

  // Color + highlight
  const colorInput = el("input");
  colorInput.type = "color";
  colorInput.value = toHex(ch.color ?? "#202124");
  colorInput.addEventListener("input", () => { markC("color"); fire(); });
  cGrp.append(field("Text color", colorInput));

  const hlWrap = el("div", "cw-se-color");
  const hlInput = el("input");
  hlInput.type = "color";
  hlInput.value = toHex(ch.highlightColor ?? "#ffff00");
  let hlOn = ch.highlightColor !== undefined;
  const hlNone = el("button", undefined, hlOn ? "Clear" : "None");
  hlNone.type = "button";
  hlInput.addEventListener("input", () => { hlOn = true; hlNone.textContent = "Clear"; markC("highlightColor"); fire(); });
  hlNone.addEventListener("click", () => { hlOn = false; hlNone.textContent = "None"; markC("highlightColor"); fire(); });
  hlWrap.append(hlInput, hlNone);
  cGrp.append(field("Highlight", hlWrap));

  const vaSel = el("select");
  option(vaSel, "", "Normal");
  option(vaSel, "super", "Superscript");
  option(vaSel, "sub", "Subscript");
  vaSel.value = ch.verticalAlign ?? "";
  vaSel.addEventListener("change", () => { markC("verticalAlign"); fire(); });
  cGrp.append(field("Position", vaSel));
  root.append(cGrp);

  // -- Paragraph (paragraph styles only) -------------------------------------
  let alignSel: HTMLSelectElement | undefined;
  let lineInput: HTMLInputElement | undefined;
  let beforeInput: HTMLInputElement | undefined;
  let afterInput: HTMLInputElement | undefined;
  let indLeft: HTMLInputElement | undefined;
  let indRight: HTMLInputElement | undefined;
  let indFirst: HTMLInputElement | undefined;
  let kwn: HTMLInputElement | undefined;
  let klt: HTMLInputElement | undefined;
  let pbb: HTMLInputElement | undefined;
  let outlineSel: HTMLSelectElement | undefined;

  if (isPara) {
    const pGrp = el("div", "cw-se-grp");
    pGrp.append(el("h3", undefined, "Paragraph"));

    alignSel = el("select");
    for (const a of ALIGNS) option(alignSel, a!, a!.charAt(0).toUpperCase() + a!.slice(1));
    alignSel.value = pa.align ?? "left";
    alignSel.addEventListener("change", () => { markP("align"); fire(); });

    lineInput = el("input");
    lineInput.type = "number"; lineInput.min = "0.5"; lineInput.max = "4"; lineInput.step = "0.05";
    lineInput.value = String(pa.lineHeight ?? 1.5);
    lineInput.addEventListener("input", () => { markP("lineHeight"); fire(); });
    pGrp.append((() => { const r = el("div", "cw-se-row"); r.append(field("Alignment", alignSel!), field("Line spacing", lineInput!)); return r; })());

    beforeInput = numberPx(pa.spaceBeforePx ?? 0, () => { markP("spaceBeforePx"); fire(); });
    afterInput = numberPx(pa.spaceAfterPx ?? 0, () => { markP("spaceAfterPx"); fire(); });
    pGrp.append((() => { const r = el("div", "cw-se-row"); r.append(field("Space before (px)", beforeInput!), field("Space after (px)", afterInput!)); return r; })());

    indLeft = numberPx(pa.indentLeftPx ?? 0, () => { markP("indentLeftPx"); fire(); });
    indRight = numberPx(pa.indentRightPx ?? 0, () => { markP("indentRightPx"); fire(); });
    indFirst = numberPx(pa.indentFirstLinePx ?? 0, () => { markP("indentFirstLinePx"); fire(); });
    pGrp.append((() => { const r = el("div", "cw-se-row"); r.append(field("Indent left (px)", indLeft!), field("Indent right (px)", indRight!), field("First line (px)", indFirst!)); return r; })());

    const checks = el("div", "cw-se-checks");
    kwn = checkbox(checks, "Keep with next paragraph", pa.keepWithNext ?? false, () => { markP("keepWithNext"); fire(); });
    klt = checkbox(checks, "Keep lines together", pa.keepLinesTogether ?? false, () => { markP("keepLinesTogether"); fire(); });
    pbb = checkbox(checks, "Page break before", pa.pageBreakBefore ?? false, () => { markP("pageBreakBefore"); fire(); });
    pGrp.append(checks);

    outlineSel = el("select");
    option(outlineSel, "", "Body text");
    for (let i = 0; i < 9; i++) option(outlineSel, String(i), `Level ${i + 1}`);
    outlineSel.value = pa.outlineLevel !== undefined ? String(pa.outlineLevel) : "";
    outlineSel.addEventListener("change", () => { markP("outlineLevel"); fire(); });
    pGrp.append(field("Outline level", outlineSel));
    root.append(pGrp);
  }

  host.append(root);

  function numberPx(value: number, on: () => void): HTMLInputElement {
    const i = el("input");
    i.type = "number"; i.step = "1";
    i.value = String(Math.round(value));
    i.addEventListener("input", on);
    return i;
  }
  function checkbox(parent: HTMLElement, label: string, checked: boolean, on: () => void): HTMLInputElement {
    const l = el("label");
    const c = el("input");
    c.type = "checkbox";
    c.checked = checked;
    c.addEventListener("change", on);
    l.append(c, document.createTextNode(label));
    parent.append(l);
    return c;
  }

  const read = (): NamedStyleSpec => {
    const char: Partial<CharStyle> = {};
    if (definedChar.has("fontFamily")) char.fontFamily = fontSel.value;
    if (definedChar.has("fontSizePx")) char.fontSizePx = Math.min(96, Math.max(6, ptToPx(Number(sizeInput.value) || 12)));
    if (definedChar.has("bold")) char.bold = tglState.bold;
    if (definedChar.has("italic")) char.italic = tglState.italic;
    if (definedChar.has("underline")) char.underline = tglState.underline;
    if (definedChar.has("strikethrough")) char.strikethrough = tglState.strikethrough;
    if (definedChar.has("color")) char.color = colorInput.value;
    if (definedChar.has("highlightColor")) char.highlightColor = hlOn ? hlInput.value : undefined;
    if (definedChar.has("verticalAlign")) char.verticalAlign = vaSel.value ? (vaSel.value as "sub" | "super") : undefined;

    const para: Partial<ParaStyle> = {};
    if (isPara) {
      if (definedPara.has("align")) para.align = alignSel!.value as ParaStyle["align"];
      if (definedPara.has("lineHeight")) para.lineHeight = Number(lineInput!.value) || 1.5;
      if (definedPara.has("spaceBeforePx")) para.spaceBeforePx = Number(beforeInput!.value) || 0;
      if (definedPara.has("spaceAfterPx")) para.spaceAfterPx = Number(afterInput!.value) || 0;
      if (definedPara.has("indentLeftPx")) para.indentLeftPx = Number(indLeft!.value) || 0;
      if (definedPara.has("indentRightPx")) para.indentRightPx = Number(indRight!.value) || 0;
      if (definedPara.has("indentFirstLinePx")) para.indentFirstLinePx = Number(indFirst!.value) || 0;
      if (definedPara.has("keepWithNext")) para.keepWithNext = kwn!.checked;
      if (definedPara.has("keepLinesTogether")) para.keepLinesTogether = klt!.checked;
      if (definedPara.has("pageBreakBefore")) para.pageBreakBefore = pbb!.checked;
      if (definedPara.has("outlineLevel") && outlineSel!.value) para.outlineLevel = Number(outlineSel!.value);
    }

    const spec: NamedStyleSpec = { name: nameInput.value, type: init.spec.type, char, para };
    if (init.spec.id) spec.id = init.spec.id;
    if (basedSel.value) spec.basedOn = basedSel.value;
    return spec;
  };

  return {
    read,
    onChange: (cb) => { onChangeCb = cb; },
    dispose: () => { root.remove(); },
  };
}
