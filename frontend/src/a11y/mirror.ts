// Layer 6: accessibility. Canvas text is invisible to assistive tech — the IME
// proxy carries role=textbox; this module keeps its accessible text in sync with
// the paragraph under the caret and provides an ARIA live region for editor
// announcements. (A full mirror with real DOM selection sync is milestone 6;
// this is the minimum honest version, shipped WITH typing, not after it.)

import type { EditorState } from "../editor/state";
import { textOfBlock } from "@cw/shared";

export interface A11yMirror {
  sync(state: EditorState): void;
  announce(message: string): void;
  destroy(): void;
}

export function createA11yMirror(proxyEl: HTMLElement): A11yMirror {
  const live = document.createElement("div");
  live.setAttribute("aria-live", "polite");
  // Pinned to the origin (top/left:0) so the out-of-flow node can't sit past the
  // page content and add ~1px of scroll height; margin:0 keeps it from nudging
  // layout. Still visually hidden (1px + clip) but present for assistive tech.
  // Uses the broadly-supported `clip: rect(...)` sr-only idiom rather than
  // `clip-path: inset(50%)`, which some older assistive tech doesn't honor.
  live.style.cssText =
    "position:absolute;top:0;left:0;width:1px;height:1px;margin:0;overflow:hidden;clip:rect(0,0,0,0);white-space:nowrap;";
  document.body.appendChild(live);

  return {
    sync(state: EditorState): void {
      const focus = state.selection?.focus;
      if (!focus) return;
      const text = textOfBlock(state.doc, focus.blockId);
      proxyEl.setAttribute("aria-label", text.length > 0 ? text : "Empty paragraph");
    },
    announce(message: string): void {
      live.textContent = message;
    },
    destroy(): void {
      live.remove();
    },
  };
}
