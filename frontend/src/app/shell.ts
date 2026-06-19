// Builds the editor's DOM skeleton (the structure index.html used to hard-code)
// under a host-provided container, returning references the app wires into.
// Mirrors the original markup: toolbar, work area (outline drawer + editor pane
// with ruler + scrolling page canvas), status bar.

export interface EditorShell {
  root: HTMLDivElement;
  toolbar: HTMLDivElement;
  outline: HTMLElement;
  ruler: HTMLDivElement;
  /** Vertical ruler down the left of the scroll area, mirroring `ruler`. */
  vruler: HTMLDivElement;
  app: HTMLDivElement;
  /** Right-docked review pane (track changes + comments) — in-flow like the
   *  outline drawer, so it shrinks the editor rather than overlaying it. */
  review: HTMLElement;
  statusbar: HTMLDivElement;
}

export function buildShell(container: HTMLElement): EditorShell {
  // Class-based (not id-based) so multiple editors can share one page — see
  // ui/styles.ts. Every structural node is keyed by a `cw-*` class the shared
  // stylesheet targets under `.wordcanvas-root`.
  const div = (cls?: string, tag: keyof HTMLElementTagNameMap = "div"): HTMLElement => {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    return e;
  };

  const root = div("wordcanvas-root") as HTMLDivElement;

  const toolbar = div("cw-toolbar") as HTMLDivElement;
  const workarea = div("cw-workarea") as HTMLDivElement;
  const outline = div("cw-outline", "aside");
  const editorpane = div("cw-editorpane") as HTMLDivElement;
  // Two rows: a top ruler-row (corner spacer + horizontal ruler) and a main row
  // (vertical ruler + scroll area). The corner keeps the horizontal ruler's left
  // edge aligned with the scroll area, so its pageLeft math needs no change.
  const rulerRow = div("cw-ruler-row") as HTMLDivElement;
  const rulerCorner = div("cw-ruler-corner") as HTMLDivElement;
  const ruler = div("cw-ruler") as HTMLDivElement;
  const mainRow = div("cw-main-row") as HTMLDivElement;
  const vruler = div("cw-vruler") as HTMLDivElement;
  const app = div("cw-app") as HTMLDivElement;
  const review = div("cw-review", "aside");
  const statusbar = div("cw-statusbar") as HTMLDivElement;

  rulerRow.append(rulerCorner, ruler);
  mainRow.append(vruler, app);
  editorpane.append(rulerRow, mainRow);
  workarea.append(outline, editorpane, review);
  root.append(toolbar, workarea, statusbar);
  container.appendChild(root);

  return { root, toolbar, outline, ruler, vruler, app, review, statusbar };
}
