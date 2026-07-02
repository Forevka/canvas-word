// Shared scaffolding for the editor's floating dialogs (Font, Paragraph, Table
// Properties, Page Layout, Field Constructor, SDT Inspector, Style Manager, TOC
// Properties): the backdrop + modal + header(title, ×) + body + foot DOM, the
// Escape/×-to-close wiring, the teardown AbortController, and the floating-panel
// behavior (makeFloatingDialog) — previously copy-pasted per dialog.
//
// The shell deliberately does NOT own any CSS: each dialog keeps its own
// stylesheet (passed through to injectCssOnce untouched), because the scaffold
// rules intentionally differ per dialog (modal widths, backdrop shade). Class
// names stay `${prefix}-backdrop/-modal/-head/-x/-body/-foot`, so the rendered
// DOM is byte-identical to the hand-rolled version this replaces.

import { injectCssOnce } from "./styles";
import { makeFloatingDialog } from "./floatingDialog";

export interface DialogShellOptions {
  /** CSS class prefix, e.g. "cw-toc" → `cw-toc-backdrop`, `cw-toc-modal`, … */
  prefix: string;
  /** injectCssOnce id for the dialog's stylesheet. */
  cssId: string;
  /** The dialog's FULL stylesheet (scaffolding + content rules), unchanged. */
  css: string;
  /** Header title (an h2). Update later via the returned `titleEl`. */
  title: string;
  /** Extra header elements between the title and the × (unit selector, badge). */
  headExtras?: HTMLElement[];
  /** Additional no-drag selector(s) beyond the × button (comma-separated). */
  extraNoDrag?: string;
  /** Called exactly once when the dialog closes (×, Escape, or close()). */
  onClose?: () => void;
}

export interface DialogShell {
  backdrop: HTMLDivElement;
  modal: HTMLDivElement;
  head: HTMLDivElement;
  titleEl: HTMLHeadingElement;
  /** The header's × button (e.g. to set a tooltip). Already wired to close(). */
  closeBtn: HTMLButtonElement;
  /** Caller appends the dialog's content here. */
  body: HTMLDivElement;
  /** Caller appends the dialog's buttons here. */
  foot: HTMLDivElement;
  /** Aborted on close — attach any extra listeners with it. */
  signal: AbortSignal;
  /** Close + tear down (idempotent). Fires onClose once. */
  close(): void;
}

/** Build, mount, and wire one floating dialog. The caller populates `body` and
 *  `foot` afterwards — every modal's width is fixed by its CSS class (not by
 *  content), so the floating placement measured here stays correct. */
export function createDialogShell(o: DialogShellOptions): DialogShell {
  injectCssOnce(o.cssId, o.css);
  const div = (cls: string): HTMLDivElement => {
    const d = document.createElement("div");
    d.className = cls;
    return d;
  };
  const backdrop = div(`${o.prefix}-backdrop`);
  const modal = div(`${o.prefix}-modal`);
  // role="dialog" + a title reference for assistive tech. Deliberately NO
  // aria-modal: these are non-blocking floating panels (the page under them
  // stays visible and interactive — see makeFloatingDialog), so claiming
  // modality would tell screen readers the rest of the page is inert.
  modal.setAttribute("role", "dialog");
  modal.setAttribute("aria-labelledby", `${o.prefix}-title`);
  // Prevent clicks inside the panel from reaching the page under the (click-
  // through) backdrop — every dialog did this.
  modal.addEventListener("mousedown", (e) => e.stopPropagation());

  const head = div(`${o.prefix}-head`);
  const titleEl = document.createElement("h2");
  titleEl.id = `${o.prefix}-title`;
  titleEl.textContent = o.title;
  const xBtn = document.createElement("button");
  xBtn.className = `${o.prefix}-x`;
  xBtn.setAttribute("aria-label", "Close");
  xBtn.textContent = "×";
  head.append(titleEl, ...(o.headExtras ?? []), xBtn);

  const body = div(`${o.prefix}-body`);
  const foot = div(`${o.prefix}-foot`);
  modal.append(head, body, foot);
  backdrop.append(modal);
  document.body.append(backdrop);

  const ac = new AbortController();
  let closed = false;
  const close = (): void => {
    if (closed) return;
    closed = true;
    backdrop.remove();
    ac.abort();
    o.onClose?.();
  };
  // Capture-phase Escape so the dialog wins over the editor's own key handling.
  window.addEventListener(
    "keydown",
    (ev: KeyboardEvent) => {
      if (ev.key === "Escape") {
        ev.preventDefault();
        ev.stopPropagation();
        close();
      }
    },
    { capture: true, signal: ac.signal },
  );
  makeFloatingDialog({
    backdrop,
    modal,
    handle: head,
    signal: ac.signal,
    noDrag: o.extraNoDrag ? `.${o.prefix}-x, ${o.extraNoDrag}` : `.${o.prefix}-x`,
  });
  xBtn.addEventListener("click", close);

  return { backdrop, modal, head, titleEl, closeBtn: xBtn, body, foot, signal: ac.signal, close };
}
