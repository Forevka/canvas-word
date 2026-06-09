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
  const div = (id?: string, tag: keyof HTMLElementTagNameMap = "div"): HTMLElement => {
    const e = document.createElement(tag);
    if (id) e.id = id;
    return e;
  };

  const root = div() as HTMLDivElement;
  root.className = "wordcanvas-root";

  const toolbar = div("toolbar") as HTMLDivElement;
  const workarea = div("workarea") as HTMLDivElement;
  const outline = div("outline", "aside");
  const editorpane = div("editorpane") as HTMLDivElement;
  const ruler = div("ruler") as HTMLDivElement;
  const app = div("app") as HTMLDivElement;
  const statusbar = div("statusbar") as HTMLDivElement;

  editorpane.append(ruler, app);
  workarea.append(outline, editorpane);
  root.append(toolbar, workarea, statusbar);
  container.appendChild(root);

  return { root, toolbar, outline, ruler, app, statusbar };
}
