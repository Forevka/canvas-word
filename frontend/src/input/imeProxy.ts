// Layer 4: the IME proxy — the piece that makes canvas-based editing honest.
//
// A visually-hidden but focusable contenteditable div, kept focused while the
// editor is active and positioned at the caret's location (so native IME
// candidate windows pop up in the right place; opacity 0 — display:none would
// break IME).
//
// - beforeinput -> translated to editor callbacks; the proxy's own DOM content
//   is irrelevant and gets cleared after composition.
// - composition events stream to the wiring layer, which renders the preview
//   as TRANSIENT model edits outside the undo stack and commits the final text
//   as one real insert on compositionend.

export interface ImeProxyDeps {
  /** View-only mode: keep the proxy focusable (so the copy/keyboard path still
   *  works) but suppress the soft keyboard and announce read-only to AT. The
   *  editor core already drops every edit, so input events here are harmless. */
  readonly?: boolean;
  onInsertText(text: string): void;
  onDeleteBackward(): void;
  onDeleteForward(): void;
  onSplitParagraph(): void;
  onPaste(payload: { html: string | null; text: string | null }): void;
  onCompositionStart(): void;
  onCompositionUpdate(data: string): void;
  onCompositionEnd(data: string): void;
}

export interface ImeProxy {
  el: HTMLDivElement;
  focus(): void;
  hasFocus(): boolean;
  /** Position the proxy at the caret (container-relative CSS pixels). */
  moveTo(left: number, top: number, height: number): void;
  destroy(): void;
}

export function createImeProxy(container: HTMLElement, deps: ImeProxyDeps): ImeProxy {
  // Touch keyboards swallow Backspace when the editable is empty (no
  // deleteContentBackward fires with nothing before the caret). On coarse-pointer
  // devices we keep a hidden SENTINEL char in the proxy so the caret always has
  // something to delete; see `seed()`. Desktop keeps the empty-proxy behavior.
  const coarse = typeof matchMedia === "function" && matchMedia("(pointer: coarse)").matches;
  const SENTINEL = " "; // non-breaking space: never collapsed (white-space:pre)

  const el = document.createElement("div");
  el.contentEditable = "true";
  el.spellcheck = false;
  el.setAttribute("role", "textbox");
  el.setAttribute("aria-multiline", "true");
  el.setAttribute("aria-label", "Document editor");
  // Soft-keyboard hints. Mobile keyboards read these off the focused element;
  // they're inert on desktop. `enterkeyhint=enter` matches insertParagraph.
  el.setAttribute("inputmode", "text");
  el.setAttribute("enterkeyhint", "enter");
  // On touch we keep a sentinel char in the field, so autocapitalize/autocorrect
  // would operate on that stale content (every letter capitalized after the
  // sentinel space, suggestions for the sentinel) — disable them there. Desktop
  // (empty proxy) keeps sentence capitalization.
  el.setAttribute("autocapitalize", coarse ? "off" : "sentences");
  el.setAttribute("autocorrect", coarse ? "off" : "on");
  // View-only: inputmode=none keeps the on-screen keyboard from opening on focus
  // while leaving the element focusable for selection/copy and keyboard nav.
  if (deps.readonly) {
    el.setAttribute("inputmode", "none");
    el.setAttribute("aria-readonly", "true");
    el.setAttribute("autocorrect", "off");
    el.setAttribute("autocapitalize", "off");
  }
  // font-size:16px keeps iOS Safari from page-zooming when the proxy gains focus
  // (it has no visible text, but Safari still applies the <16px zoom-on-focus rule).
  el.style.cssText =
    "position:absolute;left:0;top:0;width:2px;height:1em;opacity:0;overflow:hidden;" +
    "white-space:pre;outline:none;user-select:text;z-index:1;font-size:16px;";
  container.appendChild(el);

  let composing = false;

  // Re-assert the sentinel and pin the caret AFTER it, so every Backspace has a
  // char to delete (and the caret can't drift before the sentinel, which would
  // make only the first Backspace fire). No-op on desktop and during composition.
  const seed = (): void => {
    if (!coarse || composing) return;
    if (el.firstChild === null || el.textContent !== SENTINEL) el.textContent = SENTINEL;
    const node = el.firstChild;
    const sel = window.getSelection();
    if (!node || !sel) return;
    const r = document.createRange();
    r.setStart(node, (node.textContent ?? "").length);
    r.collapse(true);
    sel.removeAllRanges();
    sel.addRange(r);
  };

  const onBeforeInput = (ev: InputEvent): void => {
    // All real edits are prevented below, so the sentinel never actually changes;
    // re-pin the caret after each handled key (skipped during composition).
    if (coarse) queueMicrotask(seed);
    switch (ev.inputType) {
      case "insertText":
        ev.preventDefault();
        if (ev.data) deps.onInsertText(ev.data);
        return;
      case "insertParagraph":
        ev.preventDefault();
        deps.onSplitParagraph();
        return;
      case "insertLineBreak":
        // Shift+Enter: a SOFT break — one "\v" character inside the paragraph.
        ev.preventDefault();
        deps.onInsertText("\v");
        return;
      case "deleteContentBackward":
        if (composing) return; // IME deletions arrive via compositionupdate
        ev.preventDefault();
        deps.onDeleteBackward();
        return;
      case "deleteContentForward":
        ev.preventDefault();
        deps.onDeleteForward();
        return;
      case "insertFromPaste": {
        ev.preventDefault();
        deps.onPaste({
          html: ev.dataTransfer?.getData("text/html") || null,
          text: ev.dataTransfer?.getData("text/plain") || null,
        });
        return;
      }
      case "insertCompositionText":
        return; // not cancelable; composition events carry the truth
      default:
        ev.preventDefault(); // never let the proxy accumulate arbitrary edits
    }
  };

  const onCompositionStart = (): void => {
    composing = true;
    deps.onCompositionStart();
  };
  const onCompositionUpdate = (ev: CompositionEvent): void => {
    deps.onCompositionUpdate(ev.data ?? "");
  };
  const onCompositionEnd = (ev: CompositionEvent): void => {
    composing = false;
    deps.onCompositionEnd(ev.data ?? "");
    // Discard whatever the browser staged in the proxy; on touch, restore the
    // sentinel (re-arming Backspace) instead of leaving it empty.
    if (coarse) seed();
    else el.textContent = "";
  };

  el.addEventListener("beforeinput", onBeforeInput);
  el.addEventListener("compositionstart", onCompositionStart);
  el.addEventListener("compositionupdate", onCompositionUpdate);
  el.addEventListener("compositionend", onCompositionEnd);

  return {
    el,
    focus(): void {
      el.focus({ preventScroll: true });
      seed(); // arm Backspace on touch the moment we take focus
    },
    hasFocus(): boolean {
      return document.activeElement === el;
    },
    moveTo(left: number, top: number, height: number): void {
      el.style.left = `${left}px`;
      el.style.top = `${top}px`;
      // Never 0-height: iOS won't reliably keep the soft keyboard open for a
      // zero-height focused element.
      el.style.height = `${Math.max(height, 12)}px`;
    },
    destroy(): void {
      el.remove();
    },
  };
}
