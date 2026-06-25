// Constructor / config builder. One declarative SCHEMA drives BOTH the form on
// the left and the generated `new WordCanvas({…})` snippet below the preview, so
// there's a single source of truth for every option. Editing a control updates
// the snippet live; clicking Apply tears down the editor and re-mounts it with
// the chosen options (every option is construction-time-only, so a change can
// only take effect by re-constructing — the editor exposes a clean `destroy()`).
//
// The generated snippet is intentionally MINIMAL: it emits only the options that
// differ from the library defaults, exactly what you'd paste into your own app.

import { WordCanvas, darkCanvasTheme } from "@forevka/wordcanvas";
import type { EditMode, LoadProgress, WordCanvasOptions } from "@forevka/wordcanvas";

type Cap = "read" | "suggest" | "edit";

// ===========================================================================
// State model — mirrors WordCanvasOptions. Colors are `{ on, value }`: a color
// <input> has no "unset" state and the library's true default hexes aren't
// form-knowable, so an explicit "override?" gate keeps unset colors out of the
// snippet and preview.

interface Slot {
  on: boolean;
  value: string;
}

type ThemeColorKey =
  | "canvasBackground"
  | "grid"
  | "gridMesh"
  | "externalLink"
  | "accent"
  | "footnoteRule"
  | "formattingMark"
  | "tocLeader"
  | "imagePlaceholder"
  | "columnSeparator"
  | "caret"
  | "searchHighlight"
  | "reviewPinResolved"
  | "reviewPinStroke";

type RulerColorKey = "bg" | "content" | "line" | "label";

const VIEW_DEFAULTS = {
  ruler: true,
  verticalRuler: true,
  grid: false,
  snapToGrid: false,
  gridSpacingPx: 24,
  formattingMarks: false,
  outline: true,
  bookmarks: false,
  reviewPane: false,
  activity: false,
  toolbar: true,
  ribbonCollapsed: false,
  statusBar: true,
  exportPdf: true,
  exportDocx: true,
  zoom: 1,
};

const BEHAVIOR_DEFAULTS = {
  zoomStep: 1.1,
  zoomMin: 0.25,
  zoomMax: 5,
  indentStepPx: 36,
  gridSpacingPx: 24,
};

interface State {
  backendUrl: string;
  docId: string;
  user: { id: string; firstName: string; lastName: string };
  knownUsers: string; // textarea: one "id,first,last" per line
  readonly: boolean;
  mode: EditMode;
  allowedModes: EditMode[]; // empty = all (default)
  agentTools: { on: boolean; capabilities: Cap[]; name: string };
  view: typeof VIEW_DEFAULTS;
  theme: {
    colors: Record<ThemeColorKey, Slot>;
    pageGapPx: number;
    ruler: { colors: Record<RulerColorKey, Slot>; font: string };
  };
  overrideDefaultStyles: {
    fontFamily: string;
    fontSizePx: number;
    color: Slot;
    lineHeight: number;
    headingFontFamily: string;
  };
  behavior: typeof BEHAVIOR_DEFAULTS;
  develop: boolean;
  callbacks: { onShareLink: boolean; onSave: boolean; resolveField: boolean; customizeRibbon: boolean };
}

// Color tables: key, label, and a seed hex used only as the picker's initial
// position (the value is emitted only once its "override?" box is checked).
const THEME_COLORS: { key: ThemeColorKey; label: string; seed: string }[] = [
  { key: "canvasBackground", label: "Canvas background", seed: "#e8eaed" },
  { key: "grid", label: "Table gridlines", seed: "#cccccc" },
  { key: "gridMesh", label: "Drawing-grid mesh", seed: "#000000" },
  { key: "externalLink", label: "External link", seed: "#1a0dab" },
  { key: "accent", label: "UI accent", seed: "#1a73e8" },
  { key: "footnoteRule", label: "Footnote rule", seed: "#999999" },
  { key: "formattingMark", label: "Formatting marks", seed: "#3399ff" },
  { key: "tocLeader", label: "TOC leader", seed: "#999999" },
  { key: "imagePlaceholder", label: "Image placeholder", seed: "#cccccc" },
  { key: "columnSeparator", label: "Column separator", seed: "#cccccc" },
  { key: "caret", label: "Caret", seed: "#1a1a2e" },
  { key: "searchHighlight", label: "Search highlight", seed: "#fbbc04" },
  { key: "reviewPinResolved", label: "Resolved pin fill", seed: "#9aa0a6" },
  { key: "reviewPinStroke", label: "Comment pin stroke", seed: "#ffffff" },
];
const THEME_COLOR_KEYS = THEME_COLORS.map((c) => c.key);

const RULER_COLORS: { key: RulerColorKey; label: string; seed: string }[] = [
  { key: "bg", label: "Ruler band", seed: "#c7cdd6" },
  { key: "content", label: "Ruler content area", seed: "#ffffff" },
  { key: "line", label: "Ruler ticks", seed: "#8a8f98" },
  { key: "label", label: "Ruler labels", seed: "#605e5c" },
];
const RULER_COLOR_KEYS = RULER_COLORS.map((c) => c.key);

const MODE_ORDER: EditMode[] = ["edit", "suggest", "view"];
const CAP_ORDER: Cap[] = ["read", "suggest", "edit"];

const slot = (seed: string): Slot => ({ on: false, value: seed });

function createState(): State {
  return {
    backendUrl: "",
    docId: "",
    user: { id: "", firstName: "", lastName: "" },
    knownUsers: "",
    readonly: false,
    mode: "edit",
    allowedModes: [],
    agentTools: { on: false, capabilities: [], name: "" },
    view: { ...VIEW_DEFAULTS },
    theme: {
      colors: Object.fromEntries(THEME_COLORS.map((c) => [c.key, slot(c.seed)])) as Record<ThemeColorKey, Slot>,
      pageGapPx: 24,
      ruler: {
        colors: Object.fromEntries(RULER_COLORS.map((c) => [c.key, slot(c.seed)])) as Record<RulerColorKey, Slot>,
        font: "",
      },
    },
    overrideDefaultStyles: { fontFamily: "", fontSizePx: 16, color: slot("#202124"), lineHeight: 1.5, headingFontFamily: "" },
    behavior: { ...BEHAVIOR_DEFAULTS },
    develop: false,
    callbacks: { onShareLink: false, onSave: false, resolveField: false, customizeRibbon: false },
  };
}

let state = createState();

// ===========================================================================
// collect() — diff the state against the defaults and return ONLY the
// non-default values as a plain object. Shared by the snippet (serialized) and
// the live editor (spread into the real options).

function compactUser(u: State["user"]): Record<string, string> | null {
  const id = u.id.trim();
  const firstName = u.firstName.trim();
  const lastName = u.lastName.trim();
  if (!id && !firstName && !lastName) return null;
  return { id, firstName, lastName };
}

function parseKnownUsers(raw: string): Record<string, string>[] {
  return raw
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const [id = "", firstName = "", lastName = ""] = line.split(",").map((p) => p.trim());
      return { id, firstName, lastName };
    })
    .filter((u) => u.id);
}

function diffPlain(cur: Record<string, unknown>, def: Record<string, unknown>): Record<string, unknown> | null {
  const out: Record<string, unknown> = {};
  for (const k of Object.keys(def)) {
    if (cur[k] !== def[k]) out[k] = cur[k];
  }
  return Object.keys(out).length ? out : null;
}

function collectTheme(): Record<string, unknown> | null {
  const t: Record<string, unknown> = {};
  for (const k of THEME_COLOR_KEYS) {
    const s = state.theme.colors[k];
    if (s.on) t[k] = s.value;
  }
  if (state.theme.pageGapPx !== 24) t.pageGapPx = state.theme.pageGapPx;
  const ruler: Record<string, unknown> = {};
  for (const k of RULER_COLOR_KEYS) {
    const s = state.theme.ruler.colors[k];
    if (s.on) ruler[k] = s.value;
  }
  if (state.theme.ruler.font.trim()) ruler.font = state.theme.ruler.font.trim();
  if (Object.keys(ruler).length) t.ruler = ruler;
  return Object.keys(t).length ? t : null;
}

function collectOds(): Record<string, unknown> | null {
  const o = state.overrideDefaultStyles;
  const out: Record<string, unknown> = {};
  if (o.fontFamily.trim()) out.fontFamily = o.fontFamily.trim();
  if (o.fontSizePx !== 16) out.fontSizePx = o.fontSizePx;
  if (o.color.on) out.color = o.color.value;
  if (o.lineHeight !== 1.5) out.lineHeight = o.lineHeight;
  if (o.headingFontFamily.trim()) out.headingFontFamily = o.headingFontFamily.trim();
  return Object.keys(out).length ? out : null;
}

function collect(): Record<string, unknown> {
  const o: Record<string, unknown> = {};

  if (state.backendUrl.trim()) o.backendUrl = state.backendUrl.trim();
  if (state.docId.trim()) o.docId = state.docId.trim();
  const user = compactUser(state.user);
  if (user) o.user = user;
  const known = parseKnownUsers(state.knownUsers);
  if (known.length) o.knownUsers = known;

  if (state.readonly) o.readonly = true;
  if (state.mode !== "edit") o.mode = state.mode;
  if (state.allowedModes.length && state.allowedModes.length < MODE_ORDER.length) {
    o.allowedModes = MODE_ORDER.filter((m) => state.allowedModes.includes(m));
  }

  if (state.agentTools.on) {
    const caps = state.agentTools.capabilities;
    const name = state.agentTools.name.trim();
    const restricted = caps.length > 0 && caps.length < CAP_ORDER.length;
    if (!restricted && !name) {
      o.agentTools = true; // full tool set
    } else {
      const at: Record<string, unknown> = {};
      if (restricted) at.capabilities = CAP_ORDER.filter((c) => caps.includes(c));
      if (name) at.name = name;
      o.agentTools = at;
    }
  }

  const view = diffPlain(state.view as unknown as Record<string, unknown>, VIEW_DEFAULTS as unknown as Record<string, unknown>);
  if (view) o.view = view;
  const theme = collectTheme();
  if (theme) o.theme = theme;
  const ods = collectOds();
  if (ods) o.overrideDefaultStyles = ods;
  const behavior = diffPlain(
    state.behavior as unknown as Record<string, unknown>,
    BEHAVIOR_DEFAULTS as unknown as Record<string, unknown>,
  );
  if (behavior) o.behavior = behavior;

  if (state.develop) o.develop = true;

  return o;
}

// ===========================================================================
// Snippet generation — pretty-print the collected options as JS object source.

const CALLBACK_STUBS: { key: keyof State["callbacks"]; stub: string }[] = [
  { key: "onShareLink", stub: `onShareLink: (url, docId) => { /* show the share link in your UI */ },` },
  { key: "onSave", stub: `onSave: (e) => { /* upload e.blob to your backend instead of downloading */ },` },
  { key: "resolveField", stub: `resolveField: async (req) => { /* resolve req.name from your backend */ return ""; },` },
  { key: "customizeRibbon", stub: `customizeRibbon: (api) => { /* reorder/remove built-ins, add tabs & buttons */ },` },
];

const isIdent = (k: string): boolean => /^[A-Za-z_$][\w$]*$/.test(k);

function printValue(v: unknown, indent: string): string {
  if (typeof v === "string") return JSON.stringify(v);
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  if (Array.isArray(v)) return "[" + v.map((x) => printValue(x, indent)).join(", ") + "]";
  if (v && typeof v === "object") {
    const entries = Object.entries(v as Record<string, unknown>);
    if (!entries.length) return "{}";
    const inner = indent + "  ";
    const body = entries
      .map(([k, val]) => `${inner}${isIdent(k) ? k : JSON.stringify(k)}: ${printValue(val, inner)},`)
      .join("\n");
    return `{\n${body}\n${indent}}`;
  }
  return "undefined";
}

function buildSnippet(): string {
  const data = collect();
  const lines: string[] = ["const editor = new WordCanvas({"];
  lines.push(`  container: document.getElementById("editor"),`);
  for (const [k, v] of Object.entries(data)) {
    lines.push(`  ${k}: ${printValue(v, "  ")},`);
  }
  for (const cb of CALLBACK_STUBS) {
    if (state.callbacks[cb.key]) lines.push(`  ${cb.stub}`);
  }
  lines.push(`  onLoadProgress: ({ phase, percent }) => { /* drive a loading bar (percent: 0..1) */ },`);
  lines.push("});");
  return lines.join("\n");
}

// ===========================================================================
// Live editor — real options for the preview. The four integration callbacks
// aren't wired here (they'd only reroute exports / the share dialog and would
// muddy the preview); only onLoadProgress drives the loader. The snippet still
// documents them.

function buildOptions(container: HTMLElement, onLoadProgress: (p: LoadProgress) => void): WordCanvasOptions {
  return { container, ...collect(), onLoadProgress } as unknown as WordCanvasOptions;
}

// ===========================================================================
// Form schema — the same field list that drives the snippet drives the DOM.
// Each field carries explicit get/set closures over `state` (no stringly-typed
// paths), so reassigning `state` on Reset stays type-safe.

type Field =
  | { kind: "bool"; label: string; hint?: string; sub?: boolean; get: () => boolean; set: (v: boolean) => void }
  | { kind: "enum"; label: string; options: readonly string[]; get: () => string; set: (v: string) => void }
  | { kind: "multi"; label: string; sub?: boolean; options: readonly string[]; get: () => string[]; set: (v: string[]) => void }
  | { kind: "text"; label: string; placeholder?: string; sub?: boolean; get: () => string; set: (v: string) => void }
  | { kind: "textarea"; label: string; placeholder?: string; get: () => string; set: (v: string) => void }
  | { kind: "number"; label: string; def: number; step?: string; get: () => number; set: (v: number) => void }
  | { kind: "color"; label: string; get: () => Slot; set: (v: Slot) => void }
  | { kind: "button"; label: string; onClick: () => void };

interface Section {
  legend: string;
  intro?: string;
  fields: Field[];
}

// view helpers (typed dynamic access kept local + safe — keys are finite)
const getView = (k: keyof typeof VIEW_DEFAULTS): boolean | number => state.view[k];
const setView = (k: keyof typeof VIEW_DEFAULTS, v: boolean | number): void => {
  (state.view as Record<string, boolean | number>)[k] = v;
};
const viewBool = (k: keyof typeof VIEW_DEFAULTS, label: string): Field => ({
  kind: "bool",
  label,
  get: () => getView(k) as boolean,
  set: (v) => setView(k, v),
});
const viewNum = (k: keyof typeof VIEW_DEFAULTS, label: string, step?: string): Field => ({
  kind: "number",
  label,
  def: VIEW_DEFAULTS[k] as number,
  ...(step ? { step } : {}),
  get: () => getView(k) as number,
  set: (v) => setView(k, v),
});

const behNum = (k: keyof typeof BEHAVIOR_DEFAULTS, label: string): Field => ({
  kind: "number",
  label,
  def: BEHAVIOR_DEFAULTS[k],
  step: "any",
  get: () => state.behavior[k],
  set: (v) => {
    (state.behavior as Record<string, number>)[k] = v;
  },
});

const themeColor = (key: ThemeColorKey, label: string): Field => ({
  kind: "color",
  label,
  get: () => state.theme.colors[key],
  set: (v) => {
    state.theme.colors[key] = v;
  },
});
const rulerColor = (key: RulerColorKey, label: string): Field => ({
  kind: "color",
  label,
  get: () => state.theme.ruler.colors[key],
  set: (v) => {
    state.theme.ruler.colors[key] = v;
  },
});

function applyDarkPreset(): void {
  const dt = darkCanvasTheme;
  if (typeof dt.canvasBackground === "string") state.theme.colors.canvasBackground = { on: true, value: dt.canvasBackground };
  if (dt.ruler) {
    const r = dt.ruler;
    if (typeof r.bg === "string") state.theme.ruler.colors.bg = { on: true, value: r.bg };
    if (typeof r.content === "string") state.theme.ruler.colors.content = { on: true, value: r.content };
    if (typeof r.line === "string") state.theme.ruler.colors.line = { on: true, value: r.line };
    if (typeof r.label === "string") state.theme.ruler.colors.label = { on: true, value: r.label };
  }
  renderForm();
  regen();
  setDirty(true);
}

const SCHEMA: Section[] = [
  {
    legend: "Connection & identity",
    intro: "Provide backendUrl to go online (sync, share, collaboration). Omit it for a fully offline editor.",
    fields: [
      { kind: "text", label: "backendUrl", placeholder: "https://api.example.com", get: () => state.backendUrl, set: (v) => (state.backendUrl = v) },
      { kind: "text", label: "docId", placeholder: "open on load (online)", get: () => state.docId, set: (v) => (state.docId = v) },
      { kind: "text", label: "user.id", sub: true, get: () => state.user.id, set: (v) => (state.user.id = v) },
      { kind: "text", label: "user.firstName", sub: true, get: () => state.user.firstName, set: (v) => (state.user.firstName = v) },
      { kind: "text", label: "user.lastName", sub: true, get: () => state.user.lastName, set: (v) => (state.user.lastName = v) },
      {
        kind: "textarea",
        label: "knownUsers — @-mentionable roster (one id,first,last per line)",
        placeholder: "u1,Ada,Lovelace\nu2,Alan,Turing",
        get: () => state.knownUsers,
        set: (v) => (state.knownUsers = v),
      },
    ],
  },
  {
    legend: "Mode & access",
    fields: [
      { kind: "bool", label: "readonly — view-only viewer", get: () => state.readonly, set: (v) => (state.readonly = v) },
      { kind: "enum", label: "mode", options: MODE_ORDER, get: () => state.mode, set: (v) => (state.mode = v as EditMode) },
      {
        kind: "multi",
        label: "allowedModes (none = all)",
        options: MODE_ORDER,
        get: () => state.allowedModes,
        set: (v) => (state.allowedModes = v as EditMode[]),
      },
    ],
  },
  {
    legend: "AI agent tools (WebMCP)",
    intro: "Expose the editor to AI agents over navigator.modelContext. Capabilities + name apply only when enabled.",
    fields: [
      { kind: "bool", label: "agentTools — expose to agents", get: () => state.agentTools.on, set: (v) => (state.agentTools.on = v) },
      {
        kind: "multi",
        label: "capabilities (none = all)",
        sub: true,
        options: CAP_ORDER,
        get: () => state.agentTools.capabilities,
        set: (v) => (state.agentTools.capabilities = v as Cap[]),
      },
      { kind: "text", label: "name — tool-name prefix", sub: true, get: () => state.agentTools.name, set: (v) => (state.agentTools.name = v) },
    ],
  },
  {
    legend: "Develop mode",
    intro: "Reveal the Developer ribbon tab. Its “Inspect document tree” button opens a devtools-style panel over the parsed Document model — hover a node to highlight it on the page, click to jump to it, and read each node's JSON.",
    fields: [
      { kind: "bool", label: "develop — Document-tree inspector", get: () => state.develop, set: (v) => (state.develop = v) },
    ],
  },
  {
    legend: "View chrome (view)",
    fields: [
      viewBool("ruler", "ruler"),
      viewBool("verticalRuler", "verticalRuler"),
      viewBool("grid", "grid overlay"),
      viewBool("snapToGrid", "snapToGrid"),
      viewNum("gridSpacingPx", "gridSpacingPx", "1"),
      viewBool("formattingMarks", "formattingMarks"),
      viewBool("outline", "outline pane"),
      viewBool("bookmarks", "bookmarks pane"),
      viewBool("reviewPane", "reviewPane"),
      viewBool("activity", "activity pane"),
      viewBool("toolbar", "toolbar (ribbon)"),
      viewBool("ribbonCollapsed", "ribbonCollapsed"),
      viewBool("statusBar", "statusBar"),
      viewBool("exportPdf", "exportPdf button"),
      viewBool("exportDocx", "exportDocx button"),
      viewNum("zoom", "zoom (0.25–5)", "0.05"),
    ],
  },
  {
    legend: "Theme (chrome colors)",
    intro: "On-screen only — exported PDFs keep the built-in look. Check a color's box to override it.",
    fields: [
      { kind: "button", label: "Use darkCanvasTheme preset", onClick: applyDarkPreset },
      ...THEME_COLORS.map((c) => themeColor(c.key, c.label)),
      { kind: "number", label: "pageGapPx", def: 24, step: "1", get: () => state.theme.pageGapPx, set: (v) => (state.theme.pageGapPx = v) },
      ...RULER_COLORS.map((c) => rulerColor(c.key, c.label)),
      { kind: "text", label: "ruler.font (CSS font)", get: () => state.theme.ruler.font, set: (v) => (state.theme.ruler.font = v) },
    ],
  },
  {
    legend: "Default styles (overrideDefaultStyles)",
    intro: "Library defaults for NEW/blank docs. A loaded .docx keeps its own defaults.",
    fields: [
      { kind: "text", label: "fontFamily", placeholder: "Georgia, serif", get: () => state.overrideDefaultStyles.fontFamily, set: (v) => (state.overrideDefaultStyles.fontFamily = v) },
      { kind: "number", label: "fontSizePx", def: 16, step: "1", get: () => state.overrideDefaultStyles.fontSizePx, set: (v) => (state.overrideDefaultStyles.fontSizePx = v) },
      { kind: "color", label: "color (body text)", get: () => state.overrideDefaultStyles.color, set: (v) => (state.overrideDefaultStyles.color = v) },
      { kind: "number", label: "lineHeight", def: 1.5, step: "0.05", get: () => state.overrideDefaultStyles.lineHeight, set: (v) => (state.overrideDefaultStyles.lineHeight = v) },
      { kind: "text", label: "headingFontFamily", placeholder: "Arial, sans-serif", get: () => state.overrideDefaultStyles.headingFontFamily, set: (v) => (state.overrideDefaultStyles.headingFontFamily = v) },
    ],
  },
  {
    legend: "Behavior",
    fields: [
      behNum("zoomStep", "zoomStep"),
      behNum("zoomMin", "zoomMin"),
      behNum("zoomMax", "zoomMax"),
      behNum("indentStepPx", "indentStepPx"),
      behNum("gridSpacingPx", "gridSpacingPx"),
    ],
  },
  {
    legend: "Callbacks (snippet stubs)",
    intro: "These are function options — toggle one to add a ready-to-fill stub to the snippet (not wired into the live preview).",
    fields: [
      { kind: "bool", label: "onShareLink", get: () => state.callbacks.onShareLink, set: (v) => (state.callbacks.onShareLink = v) },
      { kind: "bool", label: "onSave", get: () => state.callbacks.onSave, set: (v) => (state.callbacks.onSave = v) },
      { kind: "bool", label: "resolveField", get: () => state.callbacks.resolveField, set: (v) => (state.callbacks.resolveField = v) },
      { kind: "bool", label: "customizeRibbon", get: () => state.callbacks.customizeRibbon, set: (v) => (state.callbacks.customizeRibbon = v) },
    ],
  },
];

// ===========================================================================
// DOM refs + rendering

const $ = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T;
const configEl = $("config");
const editorEl = $("editor");
const snippetEl = $("snippet");
const snippetbar = $("snippetbar");
const loader = $("loader");
const bar = loader.querySelector(".bar") as HTMLElement;
const label = loader.querySelector(".label") as HTMLElement;
const PHASE_LABEL: Record<LoadProgress["phase"], string> = { bundle: "Loading editor…", fonts: "Loading fonts…", ready: "Ready" };

const fieldsEl = document.createElement("div");
fieldsEl.id = "fields";
configEl.append(fieldsEl);

const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string): HTMLElementTagNameMap[K] => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  return n;
};
const checkbox = (checked: boolean): HTMLInputElement => {
  const c = el("input");
  c.type = "checkbox";
  c.checked = checked;
  return c;
};

function renderField(f: Field): HTMLElement {
  const row = el("div", "row");
  if ("sub" in f && f.sub) row.classList.add("sub");

  if (f.kind === "button") {
    const b = el("button", "preset");
    b.textContent = f.label;
    b.addEventListener("click", f.onClick);
    row.append(b);
    return row;
  }

  if (f.kind === "textarea") {
    const wrap = el("div");
    wrap.style.width = "100%";
    const lab = el("label");
    lab.textContent = f.label;
    lab.style.display = "block";
    lab.style.marginBottom = "4px";
    const ta = el("textarea");
    if (f.placeholder) ta.placeholder = f.placeholder;
    ta.value = f.get();
    ta.addEventListener("input", () => {
      f.set(ta.value);
      onChange();
    });
    wrap.append(lab, ta);
    row.append(wrap);
    return row;
  }

  const lab = el("label");
  lab.textContent = f.label;
  row.append(lab);

  switch (f.kind) {
    case "bool": {
      const c = checkbox(f.get());
      c.addEventListener("change", () => {
        f.set(c.checked);
        onChange();
      });
      row.append(c);
      break;
    }
    case "enum": {
      const sel = el("select");
      for (const opt of f.options) {
        const o = el("option");
        o.value = opt;
        o.textContent = opt;
        sel.append(o);
      }
      sel.value = f.get();
      sel.addEventListener("change", () => {
        f.set(sel.value);
        onChange();
      });
      row.append(sel);
      break;
    }
    case "multi": {
      const box = el("div", "multi");
      for (const opt of f.options) {
        const l = el("label");
        const c = checkbox(f.get().includes(opt));
        c.addEventListener("change", () => {
          const cur = new Set(f.get());
          if (c.checked) cur.add(opt);
          else cur.delete(opt);
          f.set([...cur]);
          onChange();
        });
        l.append(c, document.createTextNode(" " + opt));
        box.append(l);
      }
      row.append(box);
      break;
    }
    case "text": {
      const i = el("input");
      i.type = "text";
      if (f.placeholder) i.placeholder = f.placeholder;
      i.value = f.get();
      i.addEventListener("input", () => {
        f.set(i.value);
        onChange();
      });
      row.append(i);
      break;
    }
    case "number": {
      const i = el("input");
      i.type = "number";
      if (f.step) i.step = f.step;
      i.value = String(f.get());
      i.addEventListener("input", () => {
        const n = parseFloat(i.value);
        f.set(Number.isFinite(n) ? n : f.def);
        onChange();
      });
      row.append(i);
      break;
    }
    case "color": {
      const s = f.get();
      const chk = checkbox(s.on);
      chk.title = "override?";
      const col = el("input");
      col.type = "color";
      col.value = s.value;
      col.disabled = !s.on;
      chk.addEventListener("change", () => {
        f.set({ on: chk.checked, value: f.get().value });
        col.disabled = !chk.checked;
        onChange();
      });
      col.addEventListener("input", () => {
        f.set({ on: true, value: col.value });
        chk.checked = true;
        col.disabled = false;
        onChange();
      });
      row.append(chk, col);
      break;
    }
  }
  return row;
}

function renderForm(): void {
  fieldsEl.replaceChildren();
  for (const section of SCHEMA) {
    const fs = el("fieldset");
    const lg = el("legend");
    lg.textContent = section.legend;
    fs.append(lg);
    if (section.intro) {
      const p = el("p", "hint");
      p.textContent = section.intro;
      fs.append(p);
    }
    for (const f of section.fields) fs.append(renderField(f));
    fieldsEl.append(fs);
  }
}

// ===========================================================================
// Wiring: snippet regen + dirty tracking + the destroy()/re-mount loop

function regen(): void {
  snippetEl.textContent = buildSnippet();
}
function setDirty(d: boolean): void {
  snippetbar.classList.toggle("dirty", d);
}
function onChange(): void {
  regen();
  setDirty(true);
}

let gen = 0;
let current: WordCanvas | null = null;

function rebuild(): void {
  const my = ++gen;
  current?.destroy(); // aborts listeners, closes any collab socket, removes the editor's root
  editorEl.replaceChildren(); // defensive — the container element itself is reused
  loader.classList.remove("done");
  bar.style.width = "0%";
  label.textContent = PHASE_LABEL.bundle;
  current = new WordCanvas(
    buildOptions(editorEl, ({ phase, percent }) => {
      if (my !== gen) return; // ignore a torn-down instance resolving late
      bar.style.width = `${Math.round(percent * 100)}%`;
      label.textContent = PHASE_LABEL[phase];
      if (phase === "ready") loader.classList.add("done");
    }),
  );
  (window as unknown as { __wc?: WordCanvas }).__wc = current;
  setDirty(false);
}

$("apply").addEventListener("click", () => rebuild());
$("reset").addEventListener("click", () => {
  state = createState();
  renderForm();
  regen();
  rebuild();
});
$("copy").addEventListener("click", async () => {
  await navigator.clipboard.writeText(buildSnippet());
  const btn = $("copy") as HTMLButtonElement;
  const prev = btn.textContent;
  btn.textContent = "Copied!";
  setTimeout(() => (btn.textContent = prev), 1200);
});

// First paint: render the form, show the snippet, mount the editor with defaults.
renderForm();
regen();
rebuild();
