// Canvas layout-overlay chips — toggles that draw the layout structure on the page
// (block / line / fragment boxes, baselines, table cells, margins, page info). Global
// to the canvas (independent of the active tab); cleared when the panel closes.

import { el, type DevPanelEditor } from "./types";

const OVERLAYS: [kind: string, label: string][] = [
  ["blockBoxes", "Blocks"],
  ["lineBoxes", "Lines"],
  ["fragments", "Fragments"],
  ["baselines", "Baselines"],
  ["cells", "Cells"],
  ["margins", "Margins"],
  ["pageInfo", "Page"],
];

export interface OverlayBar {
  element: HTMLElement;
  /** Turn every overlay off (on panel close). */
  reset(): void;
}

export function createOverlayBar(editor: DevPanelEditor): OverlayBar {
  const bar = el("div", "cw-dev-overlays");
  bar.append(el("span", "cw-dev-overlabel", "Overlays"));
  const chips: HTMLElement[] = [];
  const state = new Map<string, boolean>();
  for (const [kind, label] of OVERLAYS) {
    const chip = el("span", "cw-dev-chip", label);
    chip.title = `Toggle the ${label.toLowerCase()} overlay on the canvas`;
    chip.addEventListener("click", () => {
      const on = !(state.get(kind) ?? false);
      state.set(kind, on);
      chip.classList.toggle("on", on);
      editor.setDebugOverlay(kind, on);
    });
    chips.push(chip);
    bar.append(chip);
  }
  return {
    element: bar,
    reset(): void {
      for (const [kind] of OVERLAYS) editor.setDebugOverlay(kind, false);
      state.clear();
      for (const c of chips) c.classList.remove("on");
    },
  };
}
