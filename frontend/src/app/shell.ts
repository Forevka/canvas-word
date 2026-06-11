// Builds the editor's DOM skeleton (the structure index.html used to hard-code)
// under a host-provided container, returning references the app wires into.
// Mirrors the original markup: toolbar, work area (outline drawer + editor pane
// with ruler + scrolling page canvas), status bar.

export interface EditorShell {
  root: HTMLDivElement;
  toolbar: HTMLDivElement;
  outline: HTMLElement;
  ruler: HTMLDivElement;
  app: HTMLDivElement;
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
  const ruler = div("cw-ruler") as HTMLDivElement;
  const app = div("cw-app") as HTMLDivElement;
  const statusbar = div("cw-statusbar") as HTMLDivElement;

  editorpane.append(ruler, app);
  workarea.append(outline, editorpane);
  root.append(toolbar, workarea, statusbar);
  container.appendChild(root);

  return { root, toolbar, outline, ruler, app, statusbar };
}
