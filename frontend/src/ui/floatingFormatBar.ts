// Floating format mini-toolbar — Word's selection toolbar. Appears just above a
// text selection with quick character formatting (font, size, B/I/U/S, colour,
// highlight, clear) so common edits don't require a trip to the ribbon.
//
// Pure presentation + wiring: the host (editorApp.ts) injects how to read the live
// selection geometry / current format and how to apply each edit. It mirrors the
// inline image mini-toolbar's lifecycle (a body-level `position:fixed` bar,
// refreshed on change / scroll / resize) but for a text selection.

import type { CurrentFormat } from "../index";
import { injectCssOnce } from "./styles";
import { showContextMenu, type MenuEntry } from "./contextMenu";
import { placeSelectionBar, anchorInView, type AnchorRect } from "./floatingBarPosition";

/** The edits a bar button performs — each keeps the current selection intact. */
export interface FloatingFormatBarActions {
  toggle(key: "bold" | "italic" | "underline" | "strikethrough"): void;
  setFontFamily(value: string): void;
  setFontSizePt(pt: number): void;
  setColor(color: string): void;
  setHighlight(color: string): void;
  clearHighlight(): void;
  clearFormatting(): void;
}

export interface FloatingFormatBarDeps {
  /** Viewport (client-px) anchor rect for the selection, or null when there's none. */
  anchorRect(): AnchorRect | null;
  /** Live character format at the selection — pressed state + current font/size. */
  format(): CurrentFormat;
  /** True only when a non-empty (range) selection exists — the bar shows on ranges,
   *  like Word (a bare caret shows nothing). */
  hasRangeSelection(): boolean;
  /** Font family choices for the picker ({ value, label }). */
  fonts(): { value: string; label: string }[];
  /** True when the bar must stay hidden (view-only mode, an image is selected, a
   *  modal owns the screen, …). */
  suppressed(): boolean;
  /** px→pt for the size readout (Word shows points). */
  pxToPt(px: number): number;
  /** Return caret focus to the editor after an action. */
  focus(): void;
  actions: FloatingFormatBarActions;
}

export interface FloatingFormatBar {
  /** Re-evaluate visibility, pressed state, and position. Cheap; call on every
   *  change / scroll / resize. */
  refresh(): void;
  /** Hide without tearing down (e.g. on Escape / focus loss). */
  hide(): void;
  /** Remove the bar from the DOM. */
  destroy(): void;
}

/** A small, Word-like highlight palette for the highlight dropdown, plus "None". */
const HIGHLIGHT_COLORS: { label: string; value: string }[] = [
  { label: "Yellow", value: "#ffff00" },
  { label: "Bright green", value: "#00ff00" },
  { label: "Turquoise", value: "#00ffff" },
  { label: "Pink", value: "#ff00ff" },
  { label: "Gray", value: "#bfbfbf" },
];

const SIZE_PRESETS = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];

const CSS = `
.cw-fmtbar{position:fixed;display:none;align-items:center;gap:1px;z-index:41;
  background:#fff;border:1px solid #c8c6c4;border-radius:6px;padding:3px;
  box-shadow:0 3px 12px rgba(0,0,0,.18);font:13px Arial,sans-serif;color:#323130;
  user-select:none;white-space:nowrap;}
.cw-fmtbar button{height:26px;min-width:26px;display:inline-flex;align-items:center;
  justify-content:center;gap:2px;padding:0 5px;border:1px solid transparent;border-radius:4px;
  background:transparent;cursor:pointer;color:#323130;font:13px Arial,sans-serif;}
.cw-fmtbar button:hover{background:#e1dfdd;}
.cw-fmtbar button.active{background:#cfe2ff;border-color:#9ac2ff;}
.cw-fmtbar .cw-fmtbar-font{min-width:74px;max-width:120px;justify-content:space-between;overflow:hidden;}
.cw-fmtbar .cw-fmtbar-font .lbl{overflow:hidden;text-overflow:ellipsis;}
.cw-fmtbar .cw-fmtbar-size{min-width:34px;}
.cw-fmtbar .cw-fmtbar-b{font-weight:700;}
.cw-fmtbar .cw-fmtbar-i{font-style:italic;font-family:Georgia,serif;}
.cw-fmtbar .cw-fmtbar-u{text-decoration:underline;}
.cw-fmtbar .cw-fmtbar-s{text-decoration:line-through;}
.cw-fmtbar .cw-fmtbar-hl{font-weight:700;border-bottom:3px solid #ffff00;border-radius:2px;}
.cw-fmtbar .cw-fmtbar-color{font-weight:700;border-bottom:3px solid #e00000;border-radius:2px;}
.cw-fmtbar .cw-fmtbar-caret{font-size:10px;color:#605e5c;}
.cw-fmtbar .sep{width:1px;height:18px;background:#e1dfdd;margin:0 3px;flex:0 0 auto;}
@media (prefers-color-scheme: dark){
  :root[data-theme="dark"] .cw-fmtbar{background:#2b2b2b;border-color:#4a4a4a;color:#e6e6e6;}
  :root[data-theme="dark"] .cw-fmtbar button{color:#e6e6e6;}
  :root[data-theme="dark"] .cw-fmtbar button:hover{background:#3c3c3c;}
  :root[data-theme="dark"] .cw-fmtbar button.active{background:#264b73;border-color:#3d6ba5;}
  :root[data-theme="dark"] .cw-fmtbar .sep{background:#4a4a4a;}
}`;

export function createFloatingFormatBar(deps: FloatingFormatBarDeps): FloatingFormatBar {
  injectCssOnce("cw-fmtbar-styles", CSS);

  const bar = document.createElement("div");
  bar.className = "cw-fmtbar";
  bar.setAttribute("role", "toolbar");
  bar.setAttribute("aria-label", "Text formatting");
  // Never steal the selection: buttons act on mouseup/click, so suppressing the
  // default mousedown keeps the caret where it is (same trick the ribbon uses).
  bar.addEventListener("mousedown", (e) => e.preventDefault());
  document.body.appendChild(bar);

  // Hidden native colour picker for text colour (applied on `change` so a drag is
  // one undo step), mirroring the ribbon's shared picker.
  const colorInput = document.createElement("input");
  colorInput.type = "color";
  colorInput.style.cssText = "position:fixed;left:-9999px;width:0;height:0;";
  document.body.appendChild(colorInput);
  let lastColor = "#e00000";

  const button = (cls: string, title: string, onClick: () => void): HTMLButtonElement => {
    const b = document.createElement("button");
    b.type = "button";
    if (cls) b.className = cls;
    b.title = title;
    b.addEventListener("click", onClick);
    bar.appendChild(b);
    return b;
  };
  const sep = (): void => {
    const s = document.createElement("div");
    s.className = "sep";
    bar.appendChild(s);
  };
  const caret = (): HTMLSpanElement => {
    const c = document.createElement("span");
    c.className = "cw-fmtbar-caret";
    c.textContent = "▾";
    return c;
  };
  const act = (fn: () => void): void => {
    fn();
    deps.focus();
  };
  // Reflect a toggle button's pressed state to both the visual class and assistive
  // tech (the bar is a role="toolbar", so its toggles should expose aria-pressed).
  const setPressed = (b: HTMLButtonElement, on: boolean): void => {
    b.classList.toggle("active", on);
    b.setAttribute("aria-pressed", String(on));
  };
  /** A pressable toggle button (seeds aria-pressed="false"). */
  const toggleButton = (cls: string, title: string, onClick: () => void): HTMLButtonElement => {
    const b = button(cls, title, onClick);
    b.setAttribute("aria-pressed", "false");
    return b;
  };

  // --- Font family -----------------------------------------------------------
  const fontBtn = button("cw-fmtbar-font", "Font", () => {
    const cur = deps.format().fontFamily?.toLowerCase();
    const entries: MenuEntry[] = deps.fonts().map(
      (f): MenuEntry => ({
        kind: "item",
        label: f.label,
        ...(f.value.toLowerCase() === cur ? { shortcut: "✓" } : {}),
        onClick: () => act(() => deps.actions.setFontFamily(f.value)),
      }),
    );
    const r = fontBtn.getBoundingClientRect();
    showContextMenu(r.left, r.bottom + 2, entries);
  });
  const fontLbl = document.createElement("span");
  fontLbl.className = "lbl";
  fontLbl.textContent = "Font";
  fontBtn.append(fontLbl, caret());

  sep();

  // --- Font size -------------------------------------------------------------
  const curPt = (): number => {
    const px = deps.format().fontSizePx;
    return px !== null ? deps.pxToPt(px) : 11;
  };
  button("", "Shrink font", () => act(() => deps.actions.setFontSizePt(curPt() - 1))).textContent = "A−";
  const sizeBtn = button("cw-fmtbar-size", "Font size", () => {
    const entries: MenuEntry[] = SIZE_PRESETS.map((p) => ({
      kind: "item",
      label: String(p),
      onClick: () => act(() => deps.actions.setFontSizePt(p)),
    }));
    const r = sizeBtn.getBoundingClientRect();
    showContextMenu(r.left, r.bottom + 2, entries);
  });
  button("", "Grow font", () => act(() => deps.actions.setFontSizePt(curPt() + 1))).textContent = "A+";

  sep();

  // --- Bold / Italic / Underline / Strikethrough -----------------------------
  const bBtn = toggleButton("cw-fmtbar-b", "Bold (Ctrl+B)", () => act(() => deps.actions.toggle("bold")));
  bBtn.textContent = "B";
  const iBtn = toggleButton("cw-fmtbar-i", "Italic (Ctrl+I)", () => act(() => deps.actions.toggle("italic")));
  iBtn.textContent = "I";
  const uBtn = toggleButton("cw-fmtbar-u", "Underline (Ctrl+U)", () => act(() => deps.actions.toggle("underline")));
  uBtn.textContent = "U";
  const sBtn = toggleButton("cw-fmtbar-s", "Strikethrough", () => act(() => deps.actions.toggle("strikethrough")));
  sBtn.textContent = "S";

  sep();

  // --- Text colour + Highlight ----------------------------------------------
  const colorBtn = button("cw-fmtbar-color", "Font colour", () => {
    colorInput.value = lastColor;
    colorInput.onchange = () =>
      act(() => {
        lastColor = colorInput.value;
        colorBtn.style.borderBottomColor = lastColor;
        deps.actions.setColor(lastColor);
      });
    colorInput.click();
  });
  colorBtn.textContent = "A";

  const hlBtn = toggleButton("cw-fmtbar-hl", "Text highlight colour", () => {
    const entries: MenuEntry[] = [
      ...HIGHLIGHT_COLORS.map(
        (c): MenuEntry => ({
          kind: "item",
          label: c.label,
          icon: `<span style="display:inline-block;width:12px;height:12px;border:1px solid #999;background:${c.value}"></span>`,
          onClick: () => act(() => deps.actions.setHighlight(c.value)),
        }),
      ),
      { kind: "sep" },
      { kind: "item", label: "No colour", onClick: () => act(() => deps.actions.clearHighlight()) },
    ];
    const r = hlBtn.getBoundingClientRect();
    showContextMenu(r.left, r.bottom + 2, entries);
  });
  hlBtn.append(document.createTextNode("H"), caret());

  sep();

  button("", "Clear all formatting", () => act(() => deps.actions.clearFormatting())).textContent = "⌫";

  // --- Visibility / positioning ---------------------------------------------
  let visible = false;
  const hide = (): void => {
    if (!visible) return;
    visible = false;
    bar.style.display = "none";
  };

  const refresh = (): void => {
    if (deps.suppressed() || !deps.hasRangeSelection()) return hide();
    const anchor = deps.anchorRect();
    const viewport = { width: window.innerWidth, height: window.innerHeight };
    if (!anchor || !anchorInView(anchor, viewport)) return hide();

    // Sync pressed state + labels BEFORE measuring (label widths affect layout).
    const f = deps.format();
    setPressed(bBtn, f.bold);
    setPressed(iBtn, f.italic);
    setPressed(uBtn, f.underline);
    setPressed(sBtn, f.strikethrough);
    setPressed(hlBtn, f.highlight);
    if (f.fontFamily) {
      const match = deps.fonts().find((x) => x.value.toLowerCase() === f.fontFamily!.toLowerCase());
      fontLbl.textContent = match ? match.label : f.fontFamily;
    } else {
      fontLbl.textContent = "Font";
    }
    sizeBtn.textContent = f.fontSizePx !== null ? String(deps.pxToPt(f.fontSizePx)) : "";

    bar.style.display = "flex";
    visible = true;
    const { left, top } = placeSelectionBar(
      anchor,
      { width: bar.offsetWidth, height: bar.offsetHeight },
      viewport,
    );
    bar.style.left = `${left}px`;
    bar.style.top = `${top}px`;
  };

  const destroy = (): void => {
    bar.remove();
    colorInput.remove();
  };

  return { refresh, hide, destroy };
}
