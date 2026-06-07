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
  const el = document.createElement("div");
  el.contentEditable = "true";
  el.spellcheck = false;
  el.setAttribute("role", "textbox");
  el.setAttribute("aria-multiline", "true");
  el.setAttribute("aria-label", "Document editor");
  el.style.cssText =
    "position:absolute;left:0;top:0;width:2px;height:1em;opacity:0;overflow:hidden;" +
    "white-space:pre;outline:none;user-select:text;z-index:1;";
  container.appendChild(el);

  let composing = false;

  const onBeforeInput = (ev: InputEvent): void => {
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
    el.textContent = ""; // discard whatever the browser staged in the proxy
  };

  el.addEventListener("beforeinput", onBeforeInput);
  el.addEventListener("compositionstart", onCompositionStart);
  el.addEventListener("compositionupdate", onCompositionUpdate);
  el.addEventListener("compositionend", onCompositionEnd);

  return {
    el,
    focus(): void {
      el.focus({ preventScroll: true });
    },
    hasFocus(): boolean {
      return document.activeElement === el;
    },
    moveTo(left: number, top: number, height: number): void {
      el.style.left = `${left}px`;
      el.style.top = `${top}px`;
      el.style.height = `${height}px`;
    },
    destroy(): void {
      el.remove();
    },
  };
}
