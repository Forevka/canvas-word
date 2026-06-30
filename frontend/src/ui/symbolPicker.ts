// The symbol picker — a floating, non-blocking panel for inserting a symbol-font
// glyph (OOXML w:sym). You choose a symbol font (Symbol, Wingdings, …) and click a
// glyph from the grid; a "recently used" row (persisted in localStorage) surfaces
// the last picks. On pick we hand the caller the font + the UPPER-CASE hex code
// point Word stores, which insertSymbolCmd turns into a w:sym run. Pattern mirrors
// ui/equationEditor.ts + makeFloatingDialog.

import { injectCssOnce } from "./styles";
import { makeFloatingDialog } from "./floatingDialog";

export interface SymbolPickerOptions {
  /** Called with the symbol font and the UPPER-CASE hex code point (e.g. "F0E0")
   *  when the user clicks a glyph. */
  onPick: (font: string, charHex: string) => void;
  /** Called when the panel closes (×, Escape, or a programmatic close) so the
   *  opener can drop its retained handle. */
  onClose?: () => void;
}

export interface SymbolPickerHandle {
  close(): void;
}

/** Symbol fonts offered in the dropdown. The classic Word symbol faces store their
 *  glyphs in the Private-Use range (0xF020–0xF0FF), which is the grid we render. */
const FONTS = ["Symbol", "Wingdings", "Wingdings 2", "Wingdings 3", "Webdings"];

// Private-Use code points Word uses for symbol-font glyphs (0xF020–0xF0FF).
const FIRST_CP = 0xf020;
const LAST_CP = 0xf0ff;

const RECENT_KEY = "cw-symbol-recent";
const RECENT_MAX = 16;

interface RecentSymbol {
  font: string;
  char: string;
}

const CSS = `
.cw-symp{position:fixed;z-index:1002;width:360px;max-width:94vw;background:#fff;border:1px solid #d0d5dd;border-radius:10px;box-shadow:0 12px 40px rgba(0,0,0,.22);font:13px/1.4 system-ui,sans-serif;color:#1a1a2e;display:flex;flex-direction:column;overflow:hidden;}
.cw-symp-head{display:flex;align-items:center;gap:6px;padding:9px 12px;background:#f5f7fa;border-bottom:1px solid #e6e9ee;font-weight:600;}
.cw-symp-head .t{flex:1;}
.cw-symp-x{cursor:pointer;border:none;background:none;font-size:18px;line-height:1;color:#8a9099;padding:2px 4px;border-radius:4px;}
.cw-symp-x:hover{background:#e6e9ee;color:#1a1a2e;}
.cw-symp-body{padding:12px;display:flex;flex-direction:column;gap:10px;}
.cw-symp-row{display:flex;align-items:center;gap:8px;}
.cw-symp-row label{font-size:11px;color:#8a9099;font-weight:600;}
.cw-symp-row select{flex:1;border:1px solid #d0d5dd;border-radius:6px;padding:5px 7px;font:13px system-ui,sans-serif;}
.cw-symp-grid{display:grid;grid-template-columns:repeat(8,1fr);gap:3px;max-height:260px;overflow-y:auto;padding:2px;}
.cw-symp-grid button{cursor:pointer;border:1px solid #e1e4ea;background:#fff;border-radius:5px;height:34px;font-size:18px;line-height:1;padding:0;display:flex;align-items:center;justify-content:center;}
.cw-symp-grid button:hover{background:#eef2f7;border-color:#9aa4b2;}
.cw-symp-label{font-size:11px;color:#8a9099;font-weight:600;}
.cw-symp-recent:empty{display:none;}
.cw-symp-recent-wrap:has(.cw-symp-recent:empty){display:none;}
.cw-symp-recent{display:flex;flex-wrap:wrap;gap:3px;}
.cw-symp-recent button{cursor:pointer;border:1px solid #e1e4ea;background:#fafbfc;border-radius:5px;width:34px;height:34px;font-size:18px;line-height:1;padding:0;display:flex;align-items:center;justify-content:center;}
.cw-symp-recent button:hover{background:#eef2f7;border-color:#9aa4b2;}
`;

const readRecent = (): RecentSymbol[] => {
  try {
    const raw = localStorage.getItem(RECENT_KEY);
    if (!raw) return [];
    const parsed: unknown = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.filter((s): s is RecentSymbol => typeof s?.font === "string" && typeof s?.char === "string").slice(0, RECENT_MAX);
  } catch {
    return [];
  }
};

const writeRecent = (font: string, char: string): RecentSymbol[] => {
  const next = [{ font, char }, ...readRecent().filter((s) => !(s.font === font && s.char === char))].slice(0, RECENT_MAX);
  try {
    localStorage.setItem(RECENT_KEY, JSON.stringify(next));
  } catch {
    /* storage unavailable (private mode / quota) — recents just don't persist */
  }
  return next;
};

const glyphOf = (charHex: string): string => {
  const cp = parseInt(charHex, 16);
  return Number.isFinite(cp) && cp > 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : "�";
};

export function showSymbolPicker(opts: SymbolPickerOptions): SymbolPickerHandle {
  injectCssOnce("cw-symbol-picker", CSS);
  const ac = new AbortController();

  const backdrop = document.createElement("div");
  backdrop.style.cssText = "position:fixed;inset:0;z-index:1001;";
  const modal = document.createElement("div");
  modal.className = "cw-symp";
  // Keep canvas selection/focus from being stolen when interacting with the panel.
  modal.addEventListener("mousedown", (e) => e.stopPropagation());

  const head = document.createElement("div");
  head.className = "cw-symp-head";
  head.innerHTML = `<span class="t">Insert symbol</span>`;
  const x = document.createElement("button");
  x.className = "cw-symp-x";
  x.textContent = "×";
  x.title = "Close";
  head.appendChild(x);

  const body = document.createElement("div");
  body.className = "cw-symp-body";

  // Font selector.
  const fontRow = document.createElement("div");
  fontRow.className = "cw-symp-row";
  const fontLabel = document.createElement("label");
  fontLabel.textContent = "Font";
  const fontSel = document.createElement("select");
  for (const f of FONTS) {
    const o = document.createElement("option");
    o.value = f;
    o.textContent = f;
    fontSel.appendChild(o);
  }
  fontRow.append(fontLabel, fontSel);

  // Recently used.
  const recentWrap = document.createElement("div");
  recentWrap.className = "cw-symp-recent-wrap";
  const recentLabel = document.createElement("div");
  recentLabel.className = "cw-symp-label";
  recentLabel.textContent = "Recently used";
  const recentRow = document.createElement("div");
  recentRow.className = "cw-symp-recent";
  recentWrap.append(recentLabel, recentRow);

  // Glyph grid.
  const gridLabel = document.createElement("div");
  gridLabel.className = "cw-symp-label";
  gridLabel.textContent = "Symbols";
  const grid = document.createElement("div");
  grid.className = "cw-symp-grid";

  body.append(fontRow, recentWrap, gridLabel, grid);
  modal.append(head, body);
  backdrop.appendChild(modal);
  document.body.appendChild(backdrop);

  const pick = (font: string, char: string): void => {
    opts.onPick(font, char);
    renderRecent(writeRecent(font, char));
  };

  function renderRecent(items: RecentSymbol[]): void {
    recentRow.replaceChildren();
    for (const s of items) {
      const b = document.createElement("button");
      b.textContent = glyphOf(s.char);
      b.style.fontFamily = `${s.font}, sans-serif`;
      b.title = `${s.font} ${s.char}`;
      b.addEventListener("click", () => pick(s.font, s.char));
      recentRow.appendChild(b);
    }
  }

  function renderGrid(): void {
    const font = fontSel.value;
    grid.replaceChildren();
    for (let cp = FIRST_CP; cp <= LAST_CP; cp++) {
      const char = cp.toString(16).toUpperCase();
      const b = document.createElement("button");
      b.textContent = String.fromCodePoint(cp);
      b.style.fontFamily = `${font}, sans-serif`;
      b.title = `${font} ${char}`;
      b.addEventListener("click", () => pick(font, char));
      grid.appendChild(b);
    }
  }

  fontSel.addEventListener("change", renderGrid);

  // ---- lifecycle -----------------------------------------------------------
  let closed = false;
  const close = (): void => {
    if (closed) return;
    closed = true;
    ac.abort();
    backdrop.remove();
    opts.onClose?.();
  };
  x.addEventListener("click", close);
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") close(); }, { signal: ac.signal, capture: true });

  renderRecent(readRecent());
  renderGrid();
  makeFloatingDialog({ backdrop, modal, handle: head, signal: ac.signal, noDrag: ".cw-symp-x" });

  return { close };
}
