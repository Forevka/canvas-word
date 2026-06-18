// List-style editor pane — edits a ListDefinition level by level (format, bullet
// glyph, number pattern, start, indent, hanging). Pure presentation: read()
// returns the ListDefinition; the Style Manager dispatches upsertListDefinition
// and renders the live preview ({kind:"listSample"}).

import type { ListDefinition, ListNumberFormat } from "@cw/shared";
import { injectCssOnce } from "./styles";

export interface ListStyleController {
  read(): ListDefinition;
  onChange(cb: () => void): void;
  dispose(): void;
}

export interface ListStyleEditorInit {
  def: ListDefinition;
  isNew: boolean;
}

const FORMATS: { value: ListNumberFormat; label: string }[] = [
  { value: "bullet", label: "Bullet" },
  { value: "decimal", label: "1, 2, 3" },
  { value: "lowerLetter", label: "a, b, c" },
  { value: "upperLetter", label: "A, B, C" },
  { value: "lowerRoman", label: "i, ii, iii" },
  { value: "upperRoman", label: "I, II, III" },
];

const CSS = `
.cw-le{display:flex;flex-direction:column;gap:14px;}
.cw-le-field{display:flex;flex-direction:column;gap:4px;}
.cw-le-field>label{font-size:11px;color:#5f6368;}
.cw-le-field input,.cw-le-field select{height:30px;border:1px solid #d0d4d9;border-radius:6px;padding:0 8px;font:13px Arial,sans-serif;background:#fff;}
.cw-le-row{display:flex;gap:8px;flex-wrap:wrap;}
.cw-le-row>*{flex:1 1 0;min-width:90px;}
.cw-le-h3{margin:0;font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#80868b;font-weight:600;}`;

const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string, text?: string): HTMLElementTagNameMap[K] => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
};
const field = (label: string, control: HTMLElement): HTMLElement => {
  const w = el("div", "cw-le-field");
  w.append(el("label", undefined, label), control);
  return w;
};
const option = (sel: HTMLSelectElement, value: string, label: string): void => {
  const o = el("option"); o.value = value; o.textContent = label; sel.append(o);
};

export function mountListStyleEditor(host: HTMLElement, init: ListStyleEditorInit): ListStyleController {
  injectCssOnce("cw-le-styles", CSS);
  // Deep-clone so edits don't mutate the document's definition before Apply.
  const def: ListDefinition = { id: init.def.id, levels: init.def.levels.map((l) => ({ ...l })) };
  let onChangeCb: () => void = () => {};
  const fire = (): void => onChangeCb();
  let active = 0;

  const root = el("div", "cw-le");

  // Name → id (new lists only; existing id is the docx numId key and is fixed).
  const nameInput = el("input");
  nameInput.type = "text";
  nameInput.value = def.id;
  nameInput.disabled = !init.isNew;
  nameInput.addEventListener("input", fire);
  root.append(field(init.isNew ? "List name" : "List id", nameInput));

  const levelSel = el("select");
  for (let i = 0; i < def.levels.length; i++) option(levelSel, String(i), `Level ${i + 1}`);
  levelSel.value = "0";
  root.append(field("Editing level", levelSel));

  root.append(el("h3", "cw-le-h3", "Level format"));
  const formHost = el("div", "cw-le");
  root.append(formHost);
  host.append(root);

  const buildLevelForm = (): void => {
    formHost.replaceChildren();
    const lvl = def.levels[active]!;

    const fmtSel = el("select");
    for (const f of FORMATS) option(fmtSel, f.value, f.label);
    fmtSel.value = lvl.format;
    fmtSel.addEventListener("change", () => { lvl.format = fmtSel.value as ListNumberFormat; buildLevelForm(); fire(); });
    formHost.append(field("Format", fmtSel));

    if (lvl.format === "bullet") {
      const bullet = el("input");
      bullet.type = "text"; bullet.maxLength = 4;
      bullet.value = lvl.bulletChar ?? "•";
      bullet.addEventListener("input", () => { lvl.bulletChar = bullet.value || "•"; fire(); });
      formHost.append(field("Bullet character", bullet));
    } else {
      const pat = el("input");
      pat.type = "text";
      pat.value = lvl.text || `%${active + 1}.`;
      pat.title = "%N = level N's counter, e.g. %1.%2.";
      pat.addEventListener("input", () => { lvl.text = pat.value; fire(); });
      formHost.append(field("Number pattern (%1, %2…)", pat));

      const start = el("input");
      start.type = "number"; start.min = "0"; start.step = "1";
      start.value = String(lvl.start);
      start.addEventListener("input", () => { lvl.start = Math.max(0, Number(start.value) || 1); fire(); });
      formHost.append(field("Start at", start));
    }

    const indent = el("input");
    indent.type = "number"; indent.step = "1";
    indent.value = String(Math.round(lvl.indentLeftPx));
    indent.addEventListener("input", () => { lvl.indentLeftPx = Number(indent.value) || 0; fire(); });
    const hanging = el("input");
    hanging.type = "number"; hanging.step = "1";
    hanging.value = String(Math.round(lvl.hangingPx));
    hanging.addEventListener("input", () => { lvl.hangingPx = Number(hanging.value) || 0; fire(); });
    const row = el("div", "cw-le-row");
    row.append(field("Indent (px)", indent), field("Marker hang (px)", hanging));
    formHost.append(row);
  };

  levelSel.addEventListener("change", () => { active = Number(levelSel.value); buildLevelForm(); });
  buildLevelForm();

  const slug = (s: string): string => s.trim().replace(/\s+/g, "-").replace(/[^\w-]/g, "").toLowerCase() || "list";

  return {
    read: (): ListDefinition => ({ id: init.isNew ? slug(nameInput.value) : def.id, levels: def.levels }),
    onChange: (cb) => { onChangeCb = cb; },
    dispose: () => { root.remove(); },
  };
}
