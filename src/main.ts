import { createEditor, type CurrentFormat } from "./index";
import { createLayoutEngine } from "./layout/engine";
import { sampleDoc } from "./model/sampleDoc";
import { stressDoc } from "./model/stressDoc";
import { importDocx, type ImportResult } from "./import/docx/importDocx";
import { exportDocument, type ExportFormat } from "./export/exportDocument";
import { loadEditorFonts } from "./export/shared/editorFonts";
import { TOOLBAR_FONTS } from "./fonts/clones";

// Fonts must be resolved before the first layout — pretext measures with the same
// font strings the paint layer draws with, so a late font swap would desync them.
// The editor renders the bundled metric clones (Calibri→Carlito, …) so layout
// matches the PDF/DOCX exporters exactly, with no dependency on system fonts.
await loadEditorFonts();
await document.fonts.ready;

const app = document.getElementById("app");
if (!app) throw new Error("#app not found");

const reportImport = (r: ImportResult, ms: number): void => {
  console.log(`[docx-import] blocks=${r.doc.blocks.length} warnings=${r.warnings.length} total=${ms.toFixed(1)}ms`);
  for (const w of r.warnings) console.warn(`[docx-import] ${w.code}: ${w.message}`);
};

// ?stress=<pages> loads the perf-probe document and logs layout timings.
// ?docx=<url> fetches and imports a .docx (perf probe for the import pipeline).
const params = new URLSearchParams(location.search);
const stress = params.get("stress");
const docxUrl = params.get("docx");
let doc =
  stress !== null ? stressDoc(Number(stress) || 1000)
  : docxUrl !== null ? await (async () => {
      const i0 = performance.now();
      const result = await importDocx(await (await fetch(docxUrl)).arrayBuffer(), {
        onProgress: (phase, pct) => console.log(`[docx-import] ${phase} ${(pct * 100).toFixed(0)}%`),
      });
      reportImport(result, performance.now() - i0);
      return result.doc;
    })()
  : sampleDoc();

const engine = createLayoutEngine();
const t0 = performance.now();
const tree = engine.layout(doc); // cold: every paragraph hits prepareRichInline
const t1 = performance.now();
engine.layout(doc); // warm: 100% prepare-cache hits, pure line-walk + pagination
const t2 = performance.now();
console.log(
  `[canvas-word] blocks=${doc.blocks.length} pages=${tree.pages.length} ` +
    `layout cold=${(t1 - t0).toFixed(1)}ms warm=${(t2 - t1).toFixed(1)}ms`,
);

let syncToolbar: () => void = () => {};
let syncZoom: (zoom: number) => void = () => {};
let refreshOutline: () => void = () => {};
let refreshStatus: () => void = () => {};
let refreshRuler: () => void = () => {};
let toggleRuler: () => boolean = () => false;
let refreshImageBar: () => void = () => {};
const editorOpts = {
  engine,
  onChange: () => {
    syncToolbar();
    refreshOutline();
    refreshStatus();
    refreshRuler();
    refreshImageBar();
  },
  onZoomChange: (z: number) => {
    syncZoom(z);
    refreshStatus();
    refreshRuler();
    refreshImageBar();
  },
};
let editor = createEditor(app, doc, editorOpts);

// Replace the open document (docx import): tear down and rebuild the editor —
// the layout engine is reused so its caches survive across documents.
const replaceDocument = (next: typeof doc): void => {
  editor.destroy();
  doc = next;
  editor = createEditor(app, doc, editorOpts);
  refreshOutline();
  refreshStatus();
  refreshRuler();
  window.__cw = { doc, tree: undefined, engine, editor, createLayoutEngine, sampleDoc, stressDoc };
};

const openDocxFile = async (file: File): Promise<void> => {
  const i0 = performance.now();
  try {
    const result = await importDocx(file);
    reportImport(result, performance.now() - i0);
    replaceDocument(result.doc);
  } catch (e) {
    console.error("[docx-import]", e);
    alert(`Could not open "${file.name}": ${e instanceof Error ? e.message : String(e)}`);
  }
};

// ---- demo toolbar (app chrome, not editor core) ----------------------------
import {
  insertImage,
  insertImageInCell,
  insertTable,
  insertTableRowCmd,
  insertTableColumnCmd,
  deleteTableRowCmd,
  deleteTableColumnCmd,
  deleteTableCmd,
  mergeCellsCmd,
  unmergeCellCmd,
  setImageProps,
  applyNamedStyle,
  updateStyleToSelection,
  createStyleFromSelection,
  setParaProps,
  toggleList,
  toggleMultilevelList,
  toggleHighlight,
  toggleVerticalAlign,
  changeCaseCmd,
  adjustIndentCmd,
  clearCharFormatting,
  type CaseMode,
  setLinkCmd,
  insertPageBreak,
  insertSectionBreak,
  applyPageSetup,
  pageSetupAt,
  setBandVariantEnabled,
  insertTocCmd,
  insertFootnoteCmd,
  insertContentControl,
  removeContentControl,
  sdtAtPosition,
} from "./editor/commands";
import { defaultStylesheet, resolveStyle, styleById } from "./model/stylesheet";
import { paragraphsOf, textOfRuns } from "./model/text";
import { ICONS } from "./ui/icons";

const TOOLBAR_SVG =
  "data:image/svg+xml," +
  encodeURIComponent(
    `<svg xmlns="http://www.w3.org/2000/svg" width="280" height="100">` +
      `<rect width="280" height="100" rx="8" fill="#34a853"/>` +
      `<text x="140" y="56" font-family="Arial" font-size="18" fill="#fff" text-anchor="middle">inserted image</text>` +
      `</svg>`,
  );

// One function assigned by the find-bar block below; the ribbon's Find button
// and the global Ctrl+F shortcut share it.
let openFind: () => void = () => {};

const toolbar = document.getElementById("toolbar");
if (toolbar) {
  // ===== Word-style tabbed ribbon ==========================================
  // A tab strip drives a stack of panels (one visible at a time). Each panel
  // holds labeled groups of controls. The Home tab mirrors Word's layout; the
  // app's other real commands live on Insert / Layout / Table / View. Features
  // the engine/renderer can't do yet render as disabled stubs (greyed, tooltip).
  const CARET =
    `<svg class="caret" viewBox="0 0 8 8" fill="none" stroke="currentColor" ` +
    `stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"><path d="M1.5 3 4 5.5 6.5 3"/></svg>`;
  const chevron = (up: boolean): string =>
    `<svg width="12" height="12" viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="1.4" ` +
    `stroke-linecap="round" stroke-linejoin="round"><path d="${up ? "M3 7.5 6 4.5 9 7.5" : "M3 4.5 6 7.5 9 4.5"}"/></svg>`;
  const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string): HTMLElementTagNameMap[K] => {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    return e;
  };

  const tabsBar = el("div", "rib-tabs");
  const bodies = el("div", "rib-bodies");
  toolbar.append(tabsBar, bodies);
  const tabButtons = new Map<string, HTMLButtonElement>();
  const tabPanels = new Map<string, HTMLDivElement>();
  const showTab = (id: string): void => {
    for (const [tid, b] of tabButtons) b.classList.toggle("active", tid === id);
    for (const [tid, p] of tabPanels) p.classList.toggle("active", tid === id);
    setCollapsed(false); // clicking a tab re-pins a collapsed ribbon (Word)
  };
  const tab = (id: string, label: string, kind?: "file"): HTMLDivElement => {
    const t = el("button", "rib-tab" + (kind === "file" ? " file" : ""));
    t.textContent = label;
    t.addEventListener("click", () => showTab(id));
    tabButtons.set(id, t);
    tabsBar.appendChild(t);
    const p = el("div", "rib-panel");
    p.dataset["tab"] = id;
    tabPanels.set(id, p);
    bodies.appendChild(p);
    return p;
  };

  // `controls` is the container the btn/select helpers append into; group() and
  // row() retarget it as the ribbon is built.
  let controls: HTMLElement = el("div");
  const group = (panel: HTMLElement, label: string): void => {
    const g = el("div", "rib-group");
    controls = el("div", "rib-controls");
    const l = el("div", "rib-label");
    l.textContent = label;
    g.append(controls, l);
    panel.appendChild(g);
  };
  /** A group whose controls stack in two rows (Font, Paragraph). Returns a
   *  `row()` that opens a fresh row and points the helpers at it. */
  const groupRows = (panel: HTMLElement, label: string): (() => void) => {
    const g = el("div", "rib-group");
    const rows = el("div", "rib-rows");
    const l = el("div", "rib-label");
    l.textContent = label;
    g.append(rows, l);
    panel.appendChild(g);
    return () => {
      controls = el("div", "rib-row");
      rows.appendChild(controls);
    };
  };
  const sep = (): void => {
    const d = el("div");
    d.style.cssText = "width:1px;height:18px;background:#e1dfdd;margin:0 3px;flex:0 0 auto;";
    controls.appendChild(d);
  };

  const btn = (icon: string, title: string, onClick: () => void, caret = false): HTMLButtonElement => {
    const b = el("button", "rib-btn");
    b.innerHTML = icon + (caret ? CARET : "");
    b.title = title;
    b.addEventListener("mousedown", (e) => e.preventDefault()); // keep IME-proxy focus
    b.addEventListener("click", onClick);
    controls.appendChild(b);
    return b;
  };
  /** Letter buttons (B, I, U, x²…) — Word renders these as styled text. */
  const txtBtn = (label: string, title: string, onClick: () => void, style = "", caret = false): HTMLButtonElement => {
    const b = el("button", "rib-btn");
    b.title = title;
    const s = el("span");
    s.textContent = label;
    if (style) s.style.cssText = style;
    b.appendChild(s);
    if (caret) b.insertAdjacentHTML("beforeend", CARET);
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", onClick);
    controls.appendChild(b);
    return b;
  };
  /** Large Paste-style button: icon over a caption (with optional caret). */
  const bigBtn = (icon: string, caption: string, title: string, onClick: () => void, caret = false): HTMLButtonElement => {
    const b = el("button", "rib-btn rib-big");
    b.title = title;
    b.innerHTML = icon + `<span class="big-cap">${caption}${caret ? CARET : ""}</span>`;
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", onClick);
    controls.appendChild(b);
    return b;
  };
  /** Disabled placeholder for a feature the engine/renderer doesn't support. */
  const stub = (inner: string, title: string): HTMLButtonElement => {
    const b = el("button", "rib-btn");
    b.innerHTML = inner;
    b.title = `${title} — not supported by the engine yet`;
    b.disabled = true;
    controls.appendChild(b);
    return b;
  };
  const select = (title: string, width: number): HTMLSelectElement => {
    const s = el("select");
    s.title = title;
    s.style.width = `${width}px`;
    s.addEventListener("mousedown", (e) => e.stopPropagation());
    controls.appendChild(s);
    return s;
  };
  const opt = (s: HTMLSelectElement, value: string, label: string): void => {
    const o = el("option");
    o.value = value;
    o.textContent = label;
    s.appendChild(o);
  };

  // Toggle buttons paint a pressed state from the caret's format (Word). Each
  // registers a predicate; syncToolbar re-evaluates them on every change.
  const toggleButtons: { el: HTMLButtonElement; active: (f: CurrentFormat) => boolean }[] = [];
  const toggle = (b: HTMLButtonElement, active: (f: CurrentFormat) => boolean): HTMLButtonElement => {
    toggleButtons.push({ el: b, active });
    return b;
  };

  // ---- anchored popovers (palettes, menus, pickers, dialogs) --------------
  let activePop: HTMLElement | null = null;
  const closePop = (): void => {
    if (!activePop) return;
    activePop.remove();
    activePop = null;
    document.removeEventListener("mousedown", onDocDown, true);
    document.removeEventListener("keydown", onPopKey, true);
  };
  const onDocDown = (e: MouseEvent): void => {
    if (activePop && !activePop.contains(e.target as Node)) closePop();
  };
  const onPopKey = (e: KeyboardEvent): void => {
    if (e.key === "Escape") {
      closePop();
      editor.focus();
    }
  };
  /** Open `content` in a popover anchored under `anchor`, clamped on-screen. */
  const openPop = (anchor: HTMLElement, content: HTMLElement): HTMLElement => {
    closePop();
    const pop = el("div", "cw-pop");
    pop.addEventListener("mousedown", (e) => e.preventDefault()); // keep editor focus
    pop.appendChild(content);
    document.body.appendChild(pop);
    const r = anchor.getBoundingClientRect();
    pop.style.left = `${Math.round(r.left)}px`;
    pop.style.top = `${Math.round(r.bottom + 3)}px`;
    const pr = pop.getBoundingClientRect();
    if (pr.right > window.innerWidth - 6) pop.style.left = `${Math.round(window.innerWidth - pr.width - 6)}px`;
    if (pr.bottom > window.innerHeight - 6) pop.style.top = `${Math.round(r.top - pr.height - 3)}px`;
    activePop = pop;
    setTimeout(() => {
      document.addEventListener("mousedown", onDocDown, true);
      document.addEventListener("keydown", onPopKey, true);
    }, 0);
    return pop;
  };
  /** A simple vertical menu of labelled actions; `current` marks the active row. */
  const menu = (
    items: { label: string; sample?: string; current?: boolean; onClick: () => void }[],
  ): HTMLElement => {
    const m = el("div", "cw-menu");
    for (const it of items) {
      const b = el("button");
      const check = el("span", "check");
      check.textContent = it.current ? "✓" : "";
      const lbl = el("span");
      lbl.textContent = it.label;
      if (it.sample) lbl.className = "sample";
      b.append(check, lbl);
      b.addEventListener("click", () => {
        closePop();
        it.onClick();
      });
      m.appendChild(b);
    }
    return m;
  };

  // Word colour palette: greys, then a standard-colour row set.
  const PALETTE = [
    "#000000", "#444444", "#666666", "#999999", "#cccccc", "#ffffff",
    "#c00000", "#ff0000", "#ffc000", "#ffff00", "#92d050", "#00b050",
    "#00b0f0", "#0070c0", "#002060", "#7030a0", "#e84393", "#a0522d",
  ];

  // One shared native colour picker, reused for "More colours…". We apply on
  // `change` only (not `input`) so a drag is one undo step.
  const colorPicker = el("input");
  colorPicker.type = "color";
  colorPicker.style.cssText = "position:fixed;left:-9999px;width:0;height:0;";
  document.body.appendChild(colorPicker);
  const pickColor = (initial: string, onPick: (c: string) => void): void => {
    colorPicker.value = initial;
    colorPicker.oninput = null;
    colorPicker.onchange = () => onPick(colorPicker.value);
    colorPicker.click();
  };
  /** Colour palette popover: swatch grid + an optional clear row + "More…". */
  const colorPopover = (
    anchor: HTMLElement,
    last: string,
    clearLabel: string | null,
    onPick: (c: string) => void,
    onClear: (() => void) | null,
  ): void => {
    const wrap = el("div");
    const grid = el("div", "cw-swatches");
    for (const c of PALETTE) {
      const s = el("button");
      s.style.background = c;
      s.title = c;
      s.addEventListener("click", () => {
        closePop();
        onPick(c);
      });
      grid.appendChild(s);
    }
    wrap.appendChild(grid);
    if (clearLabel && onClear) {
      const clear = el("button", "pop-action");
      clear.textContent = clearLabel;
      clear.addEventListener("click", () => {
        closePop();
        onClear();
      });
      wrap.appendChild(clear);
    }
    const more = el("button", "pop-action");
    more.textContent = "More colours…";
    more.addEventListener("click", () => {
      closePop();
      pickColor(last, onPick);
    });
    wrap.appendChild(more);
    openPop(anchor, wrap);
  };
  /** Word's split colour button: the face applies the last colour, the caret
   *  opens a palette popover. `clearLabel`/`onClear` add a "No Color"/"Automatic"
   *  row; `active` wires the face into the pressed-state sync (highlight toggles). */
  const swatch = (cfg: {
    face: string;
    initial: string;
    title: string;
    apply: (color: string) => void;
    clearLabel?: string;
    onClear?: () => void;
    active?: (f: CurrentFormat) => boolean;
  }): void => {
    let last = cfg.initial;
    const wrap = el("div");
    wrap.style.cssText = "display:flex;align-items:stretch;";
    const main = el("button", "rib-btn rib-swatch");
    main.title = cfg.title;
    main.innerHTML = `<span class="row">${cfg.face}</span><span class="bar" style="background:${last}"></span>`;
    const bar = main.querySelector(".bar") as HTMLElement;
    main.addEventListener("mousedown", (e) => e.preventDefault());
    main.addEventListener("click", () => {
      cfg.apply(last);
      editor.focus();
    });
    const more = el("button", "rib-btn");
    more.title = `${cfg.title} — choose colour`;
    more.style.cssText = "min-width:14px;padding:0;";
    more.innerHTML = CARET;
    more.addEventListener("mousedown", (e) => e.preventDefault());
    const pick = (c: string): void => {
      last = c;
      bar.style.background = c;
      cfg.apply(c);
      editor.focus();
    };
    more.addEventListener("click", () =>
      colorPopover(more, last, cfg.clearLabel ?? null, pick, cfg.onClear ? () => {
        cfg.onClear!();
        editor.focus();
      } : null),
    );
    wrap.append(main, more);
    controls.appendChild(wrap);
    if (cfg.active) toggleButtons.push({ el: main, active: cfg.active });
  };

  /** Word's table-size grid: hover to size, click to insert. */
  const tableGridPopover = (anchor: HTMLElement): void => {
    const ROWS = 8;
    const COLS = 10;
    const wrap = el("div");
    const grid = el("div", "cw-grid");
    grid.style.gridTemplateColumns = `repeat(${COLS}, 15px)`;
    const label = el("div", "cw-grid-label");
    label.textContent = "Insert table";
    const cells: HTMLElement[][] = [];
    const highlight = (R: number, C: number): void => {
      for (let r = 0; r < ROWS; r++)
        for (let c = 0; c < COLS; c++) cells[r]![c]!.classList.toggle("on", r <= R && c <= C);
    };
    for (let r = 0; r < ROWS; r++) {
      const row: HTMLElement[] = [];
      for (let c = 0; c < COLS; c++) {
        const cell = el("div", "cell");
        cell.addEventListener("mouseenter", () => {
          highlight(r, c);
          label.textContent = `${c + 1} × ${r + 1} table`;
        });
        cell.addEventListener("click", () => {
          closePop();
          editor.dispatch(insertTable(r + 1, c + 1));
          editor.focus();
        });
        row.push(cell);
        grid.appendChild(cell);
      }
      cells.push(row);
    }
    wrap.append(grid, label);
    openPop(anchor, wrap);
  };

  /** Hyperlink dialog: URL field + Apply / Remove, applied to the selection. */
  const linkDialog = (anchor: HTMLElement): void => {
    const wrap = el("div", "cw-dialog");
    const lab = el("label");
    lab.textContent = "Address";
    const input = el("input");
    input.type = "text";
    input.placeholder = "https://example.com";
    input.addEventListener("mousedown", (e) => e.stopPropagation()); // allow focus
    lab.appendChild(input);
    const row = el("div", "row");
    const remove = el("button", "danger");
    remove.textContent = "Remove";
    remove.addEventListener("click", () => {
      closePop();
      editor.dispatch(setLinkCmd(null));
      editor.focus();
    });
    const apply = el("button", "primary");
    apply.textContent = "Apply";
    const doApply = (): void => {
      const u = input.value.trim();
      closePop();
      editor.dispatch(setLinkCmd(u === "" ? null : u));
      editor.focus();
    };
    apply.addEventListener("click", doApply);
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") doApply();
      e.stopPropagation();
    });
    row.append(remove, apply);
    wrap.append(lab, row);
    openPop(anchor, wrap);
    setTimeout(() => input.focus(), 0);
  };

  const exportAs = async (format: ExportFormat): Promise<void> => {
    try {
      const { bytes, warnings } = await exportDocument(editor.getDocument(), format);
      const mime =
        format === "pdf"
          ? "application/pdf"
          : "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
      const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: mime }));
      const a = el("a");
      a.href = url;
      a.download = `document.${format}`;
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      if (warnings.length > 0) console.warn(`[export-${format}] warnings`, warnings);
    } catch (err) {
      console.error(`[export-${format}] failed`, err);
    }
  };
  const stylesheet = (): ReturnType<typeof defaultStylesheet> =>
    editor.getDocument().stylesheet ?? defaultStylesheet();

  // ===== File tab (simplified backstage: open / export / undo) =============
  const fileTab = tab("file", "File", "file");
  group(fileTab, "Open");
  bigBtn(ICONS.open, "Open<br>.docx", "Open a Word document", () => {
    const input = el("input");
    input.type = "file";
    input.accept = ".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    input.addEventListener("change", () => {
      const file = input.files?.[0];
      if (file) void openDocxFile(file);
    });
    input.click();
  });
  group(fileTab, "Export");
  txtBtn("PDF", "Export to PDF", () => void exportAs("pdf"), "font-size:11px;font-weight:600;");
  txtBtn("DOCX", "Export to .docx", () => void exportAs("docx"), "font-size:11px;font-weight:600;");
  group(fileTab, "Undo");
  btn(ICONS.undo, "Undo (Ctrl+Z)", () => editor.undo());
  btn(ICONS.redo, "Redo (Ctrl+Y)", () => editor.redo());

  // ===== Home tab ==========================================================
  const home = tab("home", "Home");

  // ---- Clipboard ----
  group(home, "Clipboard");
  bigBtn(ICONS.paste, "Paste", "Paste (Ctrl+V)", () => editor.paste(), true);
  {
    const stack = el("div");
    stack.style.cssText = "display:flex;flex-direction:column;gap:1px;";
    const prev = controls;
    controls = stack;
    btn(ICONS.cut, "Cut (Ctrl+X)", () => editor.cut());
    btn(ICONS.copy, "Copy (Ctrl+C)", () => editor.copy());
    btn(ICONS.painter, "Format painter (double-click = sticky)", () => editor.armFormatPainter(false));
    controls = prev;
    controls.appendChild(stack);
  }

  // ---- Font ----
  const fontRow = groupRows(home, "Font");
  fontRow();
  const fontSelect = select("Font family", 150);
  // Labels show the bundled clone we actually render — e.g. "Calibri (Carlito)".
  for (const f of TOOLBAR_FONTS) opt(fontSelect, f.value, f.label);
  fontSelect.addEventListener("change", () => {
    editor.setCharStyle({ fontFamily: fontSelect.value });
    editor.focus();
  });
  // The model is px-native; Word shows POINTS. Convert on read/write (96dpi:
  // 1pt = 4/3 px). The displayed value snaps to the nearest half-point.
  const pxToPt = (px: number): number => Math.round(px * 0.75 * 2) / 2;
  const ptToPx = (pt: number): number => (pt * 4) / 3;
  const sizeInput = el("input");
  sizeInput.type = "number";
  sizeInput.min = "1";
  sizeInput.max = "72";
  sizeInput.step = "0.5";
  sizeInput.title = "Font size (pt)";
  sizeInput.style.cssText = "width:46px;padding:0 4px;";
  sizeInput.addEventListener("mousedown", (e) => e.stopPropagation());
  // Editable combo: type any size, or pick a Word preset from the dropdown.
  const sizeList = el("datalist");
  sizeList.id = "cw-font-sizes";
  for (const p of [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72]) {
    const o = el("option");
    o.value = String(p);
    sizeList.appendChild(o);
  }
  document.body.appendChild(sizeList);
  sizeInput.setAttribute("list", sizeList.id);
  /** Apply a size given in POINTS (clamped to the model's 6–96px range). */
  const setSizePt = (pt: number): void => {
    if (!Number.isFinite(pt)) return;
    const px = Math.min(96, Math.max(6, ptToPx(pt)));
    editor.setCharStyle({ fontSizePx: px });
    editor.focus();
  };
  sizeInput.addEventListener("change", () => setSizePt(Number(sizeInput.value)));
  controls.appendChild(sizeInput);
  const curPt = (): number => pxToPt(editor.currentFormat().fontSizePx ?? 16);
  txtBtn("A", "Grow font", () => setSizePt(curPt() + 1), "font-size:15px;font-weight:600;");
  txtBtn("A", "Shrink font", () => setSizePt(curPt() - 1), "font-size:10px;font-weight:600;");
  const CASES: { label: string; mode: CaseMode }[] = [
    { label: "Sentence case", mode: "sentence" },
    { label: "lowercase", mode: "lower" },
    { label: "UPPERCASE", mode: "upper" },
    { label: "Capitalize Each Word", mode: "title" },
    { label: "tOGGLE cASE", mode: "toggle" },
  ];
  const caseBtn = txtBtn("Aa", "Change case", () => {}, "font-weight:600;", true);
  caseBtn.addEventListener("click", () =>
    openPop(
      caseBtn,
      menu(
        CASES.map((c) => ({
          label: c.label,
          onClick: () => {
            editor.dispatch(changeCaseCmd(c.mode));
            editor.focus();
          },
        })),
      ),
    ),
  );
  btn(ICONS.clearFormat, "Clear all formatting", () => {
    editor.dispatch(clearCharFormatting());
    editor.dispatch(applyNamedStyle("Normal"));
    editor.focus();
  });

  fontRow();
  toggle(txtBtn("B", "Bold (Ctrl+B)", () => editor.toggleStyle("bold"), "font-weight:700;"), (f) => f.bold);
  toggle(txtBtn("I", "Italic (Ctrl+I)", () => editor.toggleStyle("italic"), "font-style:italic;font-family:Georgia,serif;"), (f) => f.italic);
  toggle(txtBtn("U", "Underline (Ctrl+U)", () => editor.toggleStyle("underline"), "text-decoration:underline;"), (f) => f.underline);
  toggle(txtBtn("ab", "Strikethrough", () => editor.toggleStyle("strikethrough"), "text-decoration:line-through;"), (f) => f.strikethrough);
  toggle(txtBtn("x²", "Superscript", () => editor.dispatch(toggleVerticalAlign("super"))), (f) => f.superscript);
  toggle(txtBtn("x₂", "Subscript", () => editor.dispatch(toggleVerticalAlign("sub"))), (f) => f.subscript);
  sep();
  stub(`<span style="color:#2b579a;font-weight:700;">A</span>`, "Text effects (glow / shadow / outline)");
  // Highlight: face toggles the last colour (re-click clears); caret opens the
  // palette, where "No Color" strips the highlight.
  swatch({
    face: ICONS.highlight,
    initial: "#ffeb3b",
    title: "Text highlight colour",
    apply: (c) => editor.dispatch(toggleHighlight(c)),
    clearLabel: "No Color",
    onClear: () => editor.setCharStyle({ highlightColor: undefined }),
    active: (f) => f.highlight,
  });
  // Font colour: face applies the last colour; palette "Automatic" resets it.
  swatch({
    face: `<span style="font-weight:700;">A</span>`,
    initial: "#e00000",
    title: "Font colour",
    apply: (c) => editor.setCharStyle({ color: c }),
    clearLabel: "Automatic",
    onClear: () => editor.setCharStyle({ color: "#202124" }),
  });
  const fontLinkBtn = txtBtn("🔗", "Insert/remove hyperlink", () => {}, "font-size:12px;");
  fontLinkBtn.addEventListener("click", () => linkDialog(fontLinkBtn));

  // ---- Paragraph ----
  const paraRow = groupRows(home, "Paragraph");
  paraRow();
  toggle(btn(ICONS.bullets, "Bulleted list", () => editor.dispatch(toggleList("bullet")), true), (f) => f.listKind === "bullet");
  toggle(btn(ICONS.numbering, "Numbered list (Tab/Shift+Tab change level)", () => editor.dispatch(toggleList("decimal")), true), (f) => f.listKind === "number");
  btn(ICONS.multilevel, "Multilevel list (1, 1.1, 1.1.1 — Tab / Shift+Tab change level)", () => editor.dispatch(toggleMultilevelList()));
  sep();
  btn(ICONS.indentDecrease, "Decrease indent", () => editor.dispatch(adjustIndentCmd(-36)));
  btn(ICONS.indentIncrease, "Increase indent", () => editor.dispatch(adjustIndentCmd(36)));
  stub(ICONS.sort, "Sort");
  stub(ICONS.marks, "Show/hide formatting marks");

  paraRow();
  toggle(btn(ICONS.alignLeft, "Align left", () => editor.align("left")), (f) => f.align === "left");
  toggle(btn(ICONS.alignCenter, "Center", () => editor.align("center")), (f) => f.align === "center");
  toggle(btn(ICONS.alignRight, "Align right", () => editor.align("right")), (f) => f.align === "right");
  toggle(btn(ICONS.alignJustify, "Justify", () => editor.align("justify")), (f) => f.align === "justify");
  sep();
  const SPACINGS = [
    { v: 1, l: "1.0" },
    { v: 1.15, l: "1.15" },
    { v: 1.5, l: "1.5" },
    { v: 2, l: "2.0" },
    { v: 2.5, l: "2.5" },
    { v: 3, l: "3.0" },
  ];
  let curLineHeight: number | null = null;
  const spacingBtn = btn(ICONS.lineSpacing, "Line spacing", () => {}, true);
  spacingBtn.addEventListener("click", () =>
    openPop(
      spacingBtn,
      menu(
        SPACINGS.map((s) => ({
          label: s.l,
          current: curLineHeight !== null && Math.abs(curLineHeight - s.v) < 1e-6,
          onClick: () => {
            editor.dispatch(setParaProps({ lineHeight: s.v }));
            editor.focus();
          },
        })),
      ),
    ),
  );
  stub(ICONS.shading + CARET, "Paragraph shading");
  stub(ICONS.borders + CARET, "Paragraph borders");

  // ---- Styles (visual gallery) ----
  group(home, "Styles");
  const styleGallery = el("div", "rib-gallery");
  const styleCards = new Map<string, HTMLButtonElement>();
  const addStyleCard = (id: string, name: string): HTMLButtonElement => {
    const c = el("button", "style-card");
    c.title = name;
    c.innerHTML = `<span class="preview">AaBbCc</span><span class="name"></span>`;
    (c.querySelector(".name") as HTMLElement).textContent = name;
    // Render the preview in the style's resolved character formatting (Word's
    // gallery). Font size is damped to fit the card while still reading larger
    // for headings/title.
    const prev = c.querySelector(".preview") as HTMLElement;
    const ch = resolveStyle(stylesheet(), id).char;
    if (ch.fontFamily) prev.style.fontFamily = ch.fontFamily;
    prev.style.fontWeight = ch.bold ? "700" : "400";
    prev.style.fontStyle = ch.italic ? "italic" : "normal";
    const deco = `${ch.underline ? "underline " : ""}${ch.strikethrough ? "line-through" : ""}`.trim();
    if (deco) prev.style.textDecoration = deco;
    if (ch.color) prev.style.color = ch.color;
    if (ch.fontSizePx) prev.style.fontSize = `${Math.min(16, Math.max(11, ch.fontSizePx * 0.55))}px`;
    c.addEventListener("mousedown", (e) => e.preventDefault());
    c.addEventListener("click", () => {
      editor.dispatch(applyNamedStyle(id));
      editor.focus();
    });
    styleCards.set(id, c);
    styleGallery.appendChild(c);
    return c;
  };
  const rebuildStyleGallery = (): void => {
    styleGallery.textContent = "";
    styleCards.clear();
    for (const s of stylesheet().styles) addStyleCard(s.id, s.name);
  };
  rebuildStyleGallery();
  controls.appendChild(styleGallery);
  btn(ICONS.stylePencil, "Update current style to match selection", () => {
    const id = editor.currentFormat().styleId;
    if (id) editor.dispatch(updateStyleToSelection(id));
  });
  btn(ICONS.styleNew, "New style from selection…", () => {
    const name = prompt("New style name:");
    if (name) {
      editor.dispatch(createStyleFromSelection(name));
      rebuildStyleGallery();
      syncToolbar();
    }
  });

  // ---- Editing ----
  group(home, "Editing");
  {
    const col = el("div");
    col.style.cssText = "display:flex;flex-direction:column;gap:1px;";
    const prev = controls;
    controls = col;
    const wide = "width:100%;justify-content:flex-start;gap:4px;";
    const ed = (icon: string, label: string, title: string, onClick: () => void): void => {
      const b = btn(icon + `<span>${label}</span>`, title, onClick);
      b.style.cssText = wide;
    };
    ed(ICONS.find, "Find", "Find & replace (Ctrl+F)", () => openFind());
    ed(ICONS.replace, "Replace", "Replace (Ctrl+F)", () => openFind());
    ed(ICONS.select, "Select All", "Select all (Ctrl+A)", () => editor.selectAll());
    controls = prev;
    controls.appendChild(col);
  }

  // ===== Insert tab ========================================================
  const insert = tab("insert", "Insert");
  group(insert, "Pages");
  btn(ICONS.pageBreak, "Page break (Ctrl+Enter)", () => {
    editor.dispatch(insertPageBreak());
    editor.focus();
  });
  btn(ICONS.sectionBreak, "Section break — next page", () => {
    editor.dispatch(insertSectionBreak());
    editor.focus();
  });
  group(insert, "Tables");
  const insTableBtn = btn(ICONS.table, "Insert table", () => {}, true);
  insTableBtn.addEventListener("click", () => tableGridPopover(insTableBtn));
  group(insert, "Illustrations");
  btn(ICONS.image, "Insert image", () => {
    editor.dispatch(insertImage(TOOLBAR_SVG, 280, 100)); // top-level caret
    editor.dispatch(insertImageInCell(TOOLBAR_SVG, 280, 100)); // table-cell caret
  });
  group(insert, "Picture"); // acts on the selected image
  btn(ICONS.wrapSquare, "Wrap text around image (square)", () => {
    const id = editor.getSelectedObject();
    if (id) editor.dispatch(setImageProps(id, { wrap: "square", align: "left" }));
  });
  btn(ICONS.wrapInline, "Image in line with text (block)", () => {
    const id = editor.getSelectedObject();
    if (id) editor.dispatch(setImageProps(id, { wrap: "block", align: "center" }));
  });
  group(insert, "Links");
  const insLinkBtn = btn(ICONS.link, "Insert/remove hyperlink", () => {});
  insLinkBtn.addEventListener("click", () => linkDialog(insLinkBtn));
  group(insert, "References");
  btn(ICONS.toc, "Insert / update table of contents (Ctrl+click an entry jumps to it)", () => {
    editor.dispatch(insertTocCmd());
    editor.focus();
  });
  btn(ICONS.tocRefresh, "Recalculate TOC page numbers from the current layout", () => {
    const n = editor.recalculateToc();
    console.log(n > 0 ? `[toc] updated ${n} page number${n === 1 ? "" : "s"}` : "[toc] page numbers already current");
    editor.focus();
  });
  txtBtn("ab¹", "Insert footnote", () => {
    editor.dispatch(insertFootnoteCmd());
    editor.focus();
  }, "font-size:11px;");
  group(insert, "Controls");
  btn(ICONS.sdtText, "Rich text content control (wraps the selection)", () => {
    editor.dispatch(insertContentControl("richText", { alias: "Text" }));
    editor.focus();
  });
  btn(ICONS.sdtCheckbox, "Check box content control", () => {
    editor.dispatch(insertContentControl("checkbox", { alias: "Check Box" }));
    editor.focus();
  });
  btn(ICONS.sdtDropdown, "Drop-down list content control", () => {
    const raw = prompt("List items (comma-separated):", "Yes, No, N/A");
    if (raw === null) return;
    const listItems = raw
      .split(",")
      .map((s) => s.trim())
      .filter((s) => s.length > 0)
      .map((s) => ({ display: s, value: s }));
    editor.dispatch(insertContentControl("dropDown", { alias: "Drop-Down List", listItems }));
    editor.focus();
  });
  btn(ICONS.sdtDate, "Date picker content control", () => {
    editor.dispatch(insertContentControl("date", { alias: "Date", dateFormat: "M/d/yyyy" }));
    editor.focus();
  });
  btn(ICONS.sdtProps, "Content control properties & content (inspect the control at the caret)", () => {
    if (!editor.inspectContentControl()) alert("Place the caret inside a content control first.");
  });
  btn(ICONS.sdtRemove, "Remove the content control at the caret (keeps its text)", () => {
    const sel = editor.getSelection();
    const id = sel ? sdtAtPosition(editor.getDocument(), sel.focus) : null;
    if (id) editor.dispatch(removeContentControl(id, false));
    editor.focus();
  });

  // ===== Layout tab ========================================================
  const layout = tab("layout", "Layout");
  group(layout, "Page Setup");
  btn(ICONS.pageSetup, "Page setup (size, orientation, margins, columns — applies to the caret's section)", () => {
    pageSetupPanel.toggle();
  });

  // ===== Table tab (acts on the cell containing the caret) =================
  const tableTab = tab("table", "Table");
  group(tableTab, "Rows & Columns");
  btn(ICONS.rowAbove, "Insert row above", () => editor.dispatch(insertTableRowCmd("above")));
  btn(ICONS.rowBelow, "Insert row below", () => editor.dispatch(insertTableRowCmd("below")));
  btn(ICONS.colLeft, "Insert column left", () => editor.dispatch(insertTableColumnCmd("left")));
  btn(ICONS.colRight, "Insert column right", () => editor.dispatch(insertTableColumnCmd("right")));
  btn(ICONS.deleteRow, "Delete row", () => editor.dispatch(deleteTableRowCmd()));
  btn(ICONS.deleteCol, "Delete column", () => editor.dispatch(deleteTableColumnCmd()));
  btn(ICONS.deleteTable, "Delete table", () => editor.dispatch(deleteTableCmd()));
  group(tableTab, "Merge");
  btn(ICONS.mergeCells, "Merge cells (select across cells in one row)", () => editor.dispatch(mergeCellsCmd()));
  btn(ICONS.unmergeCells, "Unmerge cell", () => editor.dispatch(unmergeCellCmd()));

  // ===== View tab ==========================================================
  const view = tab("view", "View");
  group(view, "Show");
  let outlineToggle = (): void => {};
  const outlineBtn = btn(ICONS.outline, "Outline / navigation pane (jump to any heading)", () => outlineToggle());
  const rulerBtn = btn(ICONS.ruler, "Ruler", () => rulerBtn.classList.toggle("active", toggleRuler()));
  rulerBtn.classList.add("active"); // ruler shows by default
  stub(ICONS.marks, "Show/hide formatting marks");
  group(view, "Zoom");
  txtBtn("−", "Zoom out", () => editor.setZoom(editor.getZoom() / 1.1), "font-size:15px;");
  const zoomSel = select("Zoom level", 66);
  for (const z of [0.5, 0.75, 1, 1.25, 1.5, 2, 3]) opt(zoomSel, String(z), `${Math.round(z * 100)}%`);
  zoomSel.value = "1";
  zoomSel.addEventListener("change", () => editor.setZoom(parseFloat(zoomSel.value)));
  txtBtn("+", "Zoom in", () => editor.setZoom(editor.getZoom() * 1.1), "font-size:15px;");

  // ---- Outline drawer (left navigation pane) ------------------------------
  // Lists every Heading-styled paragraph; clicking one moves the caret there
  // and scrolls it into view. The DOM list is rebuilt only when the set of
  // headings actually changes (so typing doesn't reset the drawer's scroll);
  // the active highlight follows the caret on every change.
  const outlineEl = document.getElementById("outline");
  if (outlineEl) {
    const head = el("div", "outline-head");
    const title = el("span");
    title.textContent = "Outline";
    const closeBtn = el("button");
    closeBtn.innerHTML = "×";
    closeBtn.title = "Close";
    head.append(title, closeBtn);
    const list = el("div");
    list.id = "outline-list";
    outlineEl.append(head, list);

    // Detect a heading + its outline level. Real .docx files name their styles
    // "Heading 1" but keep an OPAQUE styleId (e.g. "Style27"), so we resolve the
    // level through the stylesheet's name/basedOn chain — mirroring the
    // importer's isHeading(). Built-in ids like "Heading1"/"Title" still match
    // directly (the sample document has no separate stylesheet entry for them).
    const labelLevel = (label: string | undefined): number | null => {
      if (!label) return null;
      const m = /(?:^|\s)heading\s*([1-9])/i.exec(label);
      if (m) return Number(m[1]);
      if (/^\s*title\s*$/i.test(label)) return 0;
      if (/(?:^|\s)heading(?:\s|$)/i.test(label)) return 1; // "Heading" with no number
      return null;
    };
    const headingLevel = (styleId: string | undefined): number | null => {
      if (!styleId) return null;
      const sheet = editor.getDocument().stylesheet ?? defaultStylesheet();
      let level = labelLevel(styleId);
      const seen = new Set<string>();
      for (
        let cur = styleById(sheet, styleId);
        cur && level === null && !seen.has(cur.id);
        cur = cur.basedOn ? styleById(sheet, cur.basedOn) : undefined
      ) {
        seen.add(cur.id);
        level = labelLevel(cur.name) ?? labelLevel(cur.id);
      }
      return level;
    };
    type Entry = { id: string; index: number; level: number };
    let entries: Entry[] = [];
    const buttons = new Map<string, HTMLButtonElement>();
    let lastSig = "\0"; // sentinel — guarantees the first build runs

    const updateActive = (): void => {
      const caret = editor.getSelection()?.focus.blockId ?? null;
      const ci = caret ? editor.getDocument().blocks.findIndex((b) => b.id === caret) : -1;
      let activeId: string | null = null;
      for (const e of entries) if (ci >= 0 && e.index <= ci) activeId = e.id; // nearest heading at/above caret
      for (const [id, b] of buttons) b.classList.toggle("active", id === activeId);
    };

    const build = (): void => {
      const collected: { entry: Entry; text: string }[] = [];
      editor.getDocument().blocks.forEach((b, index) => {
        if (b.kind !== "paragraph") return;
        const level = headingLevel(b.style.namedStyle);
        if (level === null) return;
        const text = b.runs.map((r) => r.text).join("").trim();
        if (text === "") return; // skip empty heading-styled paragraphs (structural artifacts)
        collected.push({ entry: { id: b.id, index, level }, text });
      });
      const sig = collected.map((c) => `${c.entry.id}|${c.entry.level}|${c.text}`).join("\n");
      if (sig !== lastSig) {
        lastSig = sig;
        entries = collected.map((c) => c.entry);
        list.textContent = "";
        buttons.clear();
        if (collected.length === 0) {
          const empty = el("div", "outline-empty");
          empty.textContent = "No headings yet. Apply a Heading style (Heading 1–9) to build an outline.";
          list.appendChild(empty);
        } else {
          for (const c of collected) {
            const item = el("button", "outline-item");
            item.style.paddingLeft = `${12 + c.entry.level * 14}px`;
            item.textContent = c.text;
            item.title = c.text;
            item.addEventListener("mousedown", (e) => e.preventDefault());
            item.addEventListener("click", () => editor.revealBlock(c.entry.id));
            buttons.set(c.entry.id, item);
            list.appendChild(item);
          }
        }
      }
      updateActive();
    };

    let open = false;
    const setOpen = (v: boolean): void => {
      open = v;
      outlineEl.classList.toggle("open", v);
      outlineBtn.classList.toggle("active", v);
      if (v) build();
    };
    closeBtn.addEventListener("click", () => setOpen(false));
    outlineToggle = (): void => setOpen(!open);
    refreshOutline = (): void => {
      if (open) build();
    };
    setOpen(true); // visible by default (Word opens the navigation pane on demand; we lead with it)
  }

  // Collapse / expand the ribbon body (keeps the tab strip). A pinned chevron
  // on the right of the tab strip toggles it; clicking any tab re-pins.
  const collapseBtn = el("button", "rib-tab");
  collapseBtn.style.marginLeft = "auto";
  collapseBtn.innerHTML = chevron(true);
  const setCollapsed = (v: boolean): void => {
    toolbar.classList.toggle("collapsed", v);
    collapseBtn.innerHTML = chevron(!v);
    collapseBtn.title = v ? "Pin the ribbon" : "Collapse the ribbon (Ctrl+F1)";
  };
  collapseBtn.addEventListener("click", () => setCollapsed(!toolbar.classList.contains("collapsed")));
  window.addEventListener("keydown", (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === "F1") {
      e.preventDefault();
      setCollapsed(!toolbar.classList.contains("collapsed"));
    }
  });
  setCollapsed(false);
  tabsBar.appendChild(collapseBtn);

  showTab("home");

  // ---- toolbar controls mirror the caret formatting -----------------------
  let lastStylesheet = editor.getDocument().stylesheet ?? null;
  syncToolbar = (): void => {
    // A .docx import swaps the whole stylesheet — rebuild the gallery so
    // imported style ids (Heading 1, …) resolve instead of sticking on Normal.
    const sheet = editor.getDocument().stylesheet ?? null;
    if (sheet !== lastStylesheet) {
      lastStylesheet = sheet;
      rebuildStyleGallery();
    }
    const f = editor.currentFormat();
    if (f.styleId) {
      // Paragraph references a style the gallery doesn't list (importer kept the
      // ref but the sheet lacks it) — surface it rather than lying.
      if (!styleCards.has(f.styleId)) addStyleCard(f.styleId, f.styleId);
      for (const [id, c] of styleCards) c.classList.toggle("active", id === f.styleId);
    }
    if (f.fontFamily) {
      const match = TOOLBAR_FONTS.find((x) => x.value.toLowerCase() === f.fontFamily!.toLowerCase());
      if (match) fontSelect.value = match.value;
    }
    if (f.fontSizePx !== null && document.activeElement !== sizeInput) {
      sizeInput.value = String(pxToPt(f.fontSizePx));
    }
    curLineHeight = f.lineHeight; // read by the line-spacing menu when opened
    // Pressed state for B/I/U/S, highlight, sub/super, lists, alignment.
    for (const t of toggleButtons) t.el.classList.toggle("active", t.active(f));
  };

  // Reflect the live zoom (also driven by Ctrl+wheel): snap the select to the
  // nearest preset, or insert a one-off "NN%" option for in-between values.
  syncZoom = (z: number): void => {
    const pct = `${Math.round(z * 100)}%`;
    const preset = [...zoomSel.options].find((o) => o.textContent === pct);
    [...zoomSel.options].filter((o) => o.dataset["custom"]).forEach((o) => o.remove());
    if (preset) {
      zoomSel.value = preset.value;
    } else {
      const o = el("option");
      o.value = String(z);
      o.textContent = pct;
      o.dataset["custom"] = "1";
      zoomSel.appendChild(o);
      zoomSel.value = String(z);
    }
  };
}

// ---- status bar (page count, word/character count, zoom slider) -------------
{
  const sb = document.getElementById("statusbar");
  if (sb) {
    const left = document.createElement("div");
    left.className = "sb-left";
    const right = document.createElement("div");
    right.className = "sb-right";
    sb.append(left, right);
    const item = (parent: HTMLElement): HTMLSpanElement => {
      const s = document.createElement("span");
      s.className = "sb-item";
      parent.appendChild(s);
      return s;
    };
    const pageItem = item(left);
    const wordItem = item(left);

    const zBtn = (label: string, title: string, onClick: () => void): void => {
      const b = document.createElement("button");
      b.className = "sb-btn";
      b.textContent = label;
      b.title = title;
      b.addEventListener("click", onClick);
      right.appendChild(b);
    };
    zBtn("−", "Zoom out", () => editor.setZoom(editor.getZoom() / 1.1));
    const slider = document.createElement("input");
    slider.type = "range";
    slider.min = "50";
    slider.max = "300";
    slider.step = "10";
    slider.title = "Zoom";
    slider.addEventListener("input", () => editor.setZoom(Number(slider.value) / 100));
    right.appendChild(slider);
    zBtn("+", "Zoom in", () => editor.setZoom(editor.getZoom() * 1.1));
    const zLabel = document.createElement("span");
    zLabel.className = "sb-item sb-zoom";
    right.appendChild(zLabel);

    refreshStatus = (): void => {
      const { pageCount, currentPage } = editor.getLayoutInfo();
      pageItem.textContent = `Page ${currentPage} of ${pageCount}`;
      let chars = 0;
      let words = 0;
      for (const p of paragraphsOf(editor.getDocument())) {
        const t = textOfRuns(p.runs);
        chars += t.length;
        const m = t.match(/\S+/g);
        if (m) words += m.length;
      }
      wordItem.textContent = `${words} word${words === 1 ? "" : "s"} · ${chars} character${chars === 1 ? "" : "s"}`;
      const pct = Math.round(editor.getZoom() * 100);
      slider.value = String(Math.min(300, Math.max(50, pct)));
      zLabel.textContent = `${pct}%`;
    };
    refreshStatus();
  }
}

// ---- horizontal ruler (inch ticks, margin shading, draggable indents) -------
{
  const ruler = document.getElementById("ruler");
  if (ruler) {
    const canvas = document.createElement("canvas");
    const leftMarker = document.createElement("div");
    leftMarker.className = "ruler-marker ruler-left";
    leftMarker.title = "Left indent — drag to set";
    const firstMarker = document.createElement("div");
    firstMarker.className = "ruler-marker ruler-first";
    firstMarker.title = "First-line indent — drag to set";
    ruler.append(canvas, leftMarker, firstMarker);

    // Geometry captured each refresh, read by the drag handlers.
    let geom: { contentLeft: number; zoom: number; paraId: string | null; li: number; fi: number } | null = null;

    const draw = (W: number, H: number, pageLeft: number, pageW: number, cL: number, cR: number, zoom: number): void => {
      const ctx = canvas.getContext("2d");
      if (!ctx) return;
      const dpr = window.devicePixelRatio || 1;
      canvas.width = Math.round(W * dpr);
      canvas.height = Math.round(H * dpr);
      canvas.style.width = `${W}px`;
      canvas.style.height = `${H}px`;
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      ctx.clearRect(0, 0, W, H);
      // page band (margins greyed, content white)
      ctx.fillStyle = "#c7cdd6";
      ctx.fillRect(pageLeft, 4, pageW, H - 8);
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(cL, 4, cR - cL, H - 8);
      // inch ticks + numbers, measured from the left content edge (Word's 0)
      const inch = 96 * zoom;
      ctx.strokeStyle = "#8a8f98";
      ctx.fillStyle = "#605e5c";
      ctx.font = "9px 'Segoe UI', sans-serif";
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.lineWidth = 1;
      for (let i = 1; cL + i * inch < cR - 2; i++) {
        const x = Math.round(cL + i * inch) + 0.5;
        ctx.beginPath();
        ctx.moveTo(x, H / 2 - 3);
        ctx.lineTo(x, H / 2 + 3);
        ctx.stroke();
        ctx.fillText(String(i), x, H / 2);
        // mid tick at the half-inch
        const hx = Math.round(cL + (i - 0.5) * inch) + 0.5;
        ctx.beginPath();
        ctx.moveTo(hx, H / 2 - 1.5);
        ctx.lineTo(hx, H / 2 + 1.5);
        ctx.stroke();
      }
    };

    refreshRuler = (): void => {
      if (ruler.classList.contains("hidden")) return;
      const ph = app.querySelector<HTMLElement>("[data-page]");
      if (!ph) {
        leftMarker.style.display = firstMarker.style.display = "none";
        return;
      }
      const appRect = app.getBoundingClientRect();
      const phRect = ph.getBoundingClientRect();
      const zoom = editor.getZoom();
      const sec = editor.getDocument().section;
      const pageLeft = phRect.left - appRect.left;
      const pageW = phRect.width;
      const cL = pageLeft + sec.marginPx.left * zoom;
      const cR = pageLeft + pageW - sec.marginPx.right * zoom;
      draw(ruler.clientWidth, ruler.clientHeight, pageLeft, pageW, cL, cR, zoom);
      // indent markers for the caret paragraph
      const sel = editor.getSelection();
      const para = sel ? paragraphsOf(editor.getDocument()).find((p) => p.id === sel.focus.blockId) : null;
      if (para) {
        const li = para.style.indentLeftPx ?? 0;
        const fi = para.style.indentFirstLinePx ?? 0;
        leftMarker.style.left = `${cL + li * zoom - 5}px`;
        firstMarker.style.left = `${cL + (li + fi) * zoom - 5}px`;
        leftMarker.style.display = firstMarker.style.display = "block";
        geom = { contentLeft: cL, zoom, paraId: para.id, li, fi };
      } else {
        leftMarker.style.display = firstMarker.style.display = "none";
        geom = null;
      }
    };

    const startDrag = (which: "left" | "first") => (e: PointerEvent): void => {
      const g = geom;
      if (!g || !g.paraId) return;
      e.preventDefault();
      let liNext = g.li;
      let fiNext = g.fi;
      const onMove = (ev: PointerEvent): void => {
        const x = ev.clientX - app.getBoundingClientRect().left;
        const px = Math.max(0, (x - g.contentLeft) / g.zoom);
        if (which === "left") {
          liNext = px;
          leftMarker.style.left = `${g.contentLeft + liNext * g.zoom - 5}px`;
          firstMarker.style.left = `${g.contentLeft + (liNext + fiNext) * g.zoom - 5}px`;
        } else {
          fiNext = px - liNext; // first line sits at left-indent + first-line-indent
          firstMarker.style.left = `${g.contentLeft + px * g.zoom - 5}px`;
        }
      };
      const onUp = (): void => {
        window.removeEventListener("pointermove", onMove);
        window.removeEventListener("pointerup", onUp);
        editor.dispatch(
          setParaProps(which === "left" ? { indentLeftPx: Math.round(liNext) } : { indentFirstLinePx: Math.round(fiNext) }),
        );
        editor.focus();
      };
      window.addEventListener("pointermove", onMove);
      window.addEventListener("pointerup", onUp);
    };
    leftMarker.addEventListener("pointerdown", startDrag("left"));
    firstMarker.addEventListener("pointerdown", startDrag("first"));

    app.addEventListener("scroll", () => refreshRuler());
    window.addEventListener("resize", () => refreshRuler());
    // expose a toggle for the View-tab button
    toggleRuler = (): boolean => {
      ruler.classList.toggle("hidden");
      const shown = !ruler.classList.contains("hidden");
      if (shown) refreshRuler();
      return shown;
    };
    refreshRuler();
  }
}

// ---- floating image mini-toolbar -------------------------------------------
// Appears above a selected image (Word's hover bar): wrap, align, delete.
{
  const bar = document.createElement("div");
  bar.id = "img-toolbar";
  document.body.appendChild(bar);
  const ibtn = (icon: string, title: string, onClick: () => void, cls = ""): void => {
    const b = document.createElement("button");
    if (cls) b.className = cls;
    b.innerHTML = icon;
    b.title = title;
    b.addEventListener("mousedown", (e) => e.preventDefault()); // keep the image selected
    b.addEventListener("click", () => {
      onClick();
      refreshImageBar();
    });
    bar.appendChild(b);
  };
  const sep = (): void => {
    const s = document.createElement("div");
    s.className = "sep";
    bar.appendChild(s);
  };
  const withImg = (fn: (id: string) => void): void => {
    const id = editor.getSelectedObject();
    if (id) fn(id);
  };
  ibtn(ICONS.wrapInline, "In line with text", () => withImg((id) => editor.dispatch(setImageProps(id, { wrap: "block", align: "center" }))));
  ibtn(ICONS.wrapSquare, "Wrap text (square)", () => withImg((id) => editor.dispatch(setImageProps(id, { wrap: "square", align: "left" }))));
  sep();
  ibtn(ICONS.alignLeft, "Align left", () => editor.align("left"));
  ibtn(ICONS.alignCenter, "Align center", () => editor.align("center"));
  ibtn(ICONS.alignRight, "Align right", () => editor.align("right"));
  sep();
  ibtn(ICONS.trash, "Delete image (Del)", () => {
    editor.deleteSelectedObject();
    editor.focus();
  }, "danger");

  refreshImageBar = (): void => {
    const r = editor.getSelectedObject() ? editor.getSelectedObjectRect() : null;
    if (!r) {
      bar.style.display = "none";
      return;
    }
    bar.style.display = "flex";
    const bw = bar.offsetWidth;
    const bh = bar.offsetHeight;
    const left = Math.max(8, Math.min(window.innerWidth - bw - 8, r.left + r.width / 2 - bw / 2));
    const top = r.top - bh - 8 < 56 ? r.top + r.height + 8 : r.top - bh - 8;
    bar.style.left = `${Math.round(left)}px`;
    bar.style.top = `${Math.round(top)}px`;
  };
  app.addEventListener("scroll", () => refreshImageBar());
  window.addEventListener("resize", () => refreshImageBar());
}

// ---- page setup panel (📐) ---------------------------------------------------
// Applies to the CARET's section: size preset, orientation, margin preset.
const pageSetupPanel = (() => {
  const SIZES: Record<string, { w: number; h: number }> = {
    Letter: { w: 816, h: 1056 }, // 8.5×11in @96dpi
    A4: { w: 794, h: 1123 },
    Legal: { w: 816, h: 1344 },
  };
  const MARGINS: Record<string, { top: number; right: number; bottom: number; left: number }> = {
    Normal: { top: 96, right: 96, bottom: 96, left: 96 },
    Narrow: { top: 48, right: 48, bottom: 48, left: 48 },
    Wide: { top: 96, right: 144, bottom: 96, left: 144 },
  };
  const panel = document.createElement("div");
  panel.style.cssText =
    "position:fixed;top:46px;right:24px;display:none;flex-direction:column;gap:6px;" +
    "background:#fff;border:1px solid #dadce0;border-radius:8px;padding:10px 12px;" +
    "box-shadow:0 2px 8px rgba(0,0,0,.18);z-index:10;font-size:13px;";
  const row = (label: string, control: HTMLElement): void => {
    const r = document.createElement("div");
    r.style.cssText = "display:flex;align-items:center;gap:8px;justify-content:space-between;";
    const l = document.createElement("span");
    l.textContent = label;
    r.append(l, control);
    panel.appendChild(r);
  };
  const mkSelect = (entries: string[]): HTMLSelectElement => {
    const s = document.createElement("select");
    s.style.cssText = "height:26px;border:1px solid #dadce0;border-radius:4px;background:#fff;font-size:13px;width:120px;";
    for (const e of entries) {
      const o = document.createElement("option");
      o.value = e;
      o.textContent = e;
      s.appendChild(o);
    }
    return s;
  };
  const sizeSel = mkSelect(Object.keys(SIZES));
  const orientSel = mkSelect(["Portrait", "Landscape"]);
  const marginSel = mkSelect(Object.keys(MARGINS));
  const colsSel = mkSelect(["One", "Two", "Three"]);
  const startInput = document.createElement("input");
  startInput.type = "number";
  startInput.min = "0";
  startInput.placeholder = "continue";
  startInput.title = "Restart page numbering at this section (blank = continue)";
  startInput.style.cssText = "height:24px;border:1px solid #dadce0;border-radius:4px;padding:0 6px;font-size:13px;width:108px;";
  const mkCheck = (): HTMLInputElement => {
    const c = document.createElement("input");
    c.type = "checkbox";
    c.style.cssText = "width:16px;height:16px;";
    return c;
  };
  const firstCheck = mkCheck();
  const evenCheck = mkCheck();
  row("Size", sizeSel);
  row("Orientation", orientSel);
  row("Margins", marginSel);
  row("Columns", colsSel);
  row("Number from", startInput);
  row("Different first page", firstCheck);
  row("Different odd & even", evenCheck);
  // Variant toggles apply immediately (independent of the geometry Apply).
  firstCheck.addEventListener("change", () => {
    editor.dispatch(setBandVariantEnabled("first", firstCheck.checked));
  });
  evenCheck.addEventListener("change", () => {
    editor.dispatch(setBandVariantEnabled("even", evenCheck.checked));
  });
  const apply = document.createElement("button");
  apply.textContent = "Apply to this section";
  apply.style.cssText = "height:28px;border:1px solid #1a73e8;border-radius:4px;background:#1a73e8;color:#fff;cursor:pointer;font-size:13px;";
  apply.addEventListener("click", () => {
    const size = SIZES[sizeSel.value]!;
    const landscape = orientSel.value === "Landscape";
    const colCount = { One: 1, Two: 2, Three: 3 }[colsSel.value] ?? 1;
    const start = startInput.value.trim() === "" ? null : Math.max(0, Number(startInput.value));
    editor.dispatch(
      applyPageSetup({
        pageWidthPx: landscape ? size.h : size.w,
        pageHeightPx: landscape ? size.w : size.h,
        marginPx: { ...MARGINS[marginSel.value]! },
        columns: colCount > 1 ? { count: colCount, gapPx: 24 } : null,
        pageNumberStart: start !== null && Number.isFinite(start) ? start : null,
      }),
    );
    panel.style.display = "none";
    editor.focus();
  });
  panel.appendChild(apply);
  document.body.appendChild(panel);

  return {
    toggle(): void {
      if (panel.style.display === "flex") {
        panel.style.display = "none";
        return;
      }
      // Seed controls from the caret's section.
      const geo = pageSetupAt({ doc: editor.getDocument(), selection: editor.getSelection() });
      const landscape = geo.pageWidthPx > geo.pageHeightPx;
      const w = landscape ? geo.pageHeightPx : geo.pageWidthPx;
      const h = landscape ? geo.pageWidthPx : geo.pageHeightPx;
      sizeSel.value =
        Object.keys(SIZES).find((k) => SIZES[k]!.w === w && SIZES[k]!.h === h) ?? "Letter";
      orientSel.value = landscape ? "Landscape" : "Portrait";
      marginSel.value =
        Object.keys(MARGINS).find((k) => {
          const m = MARGINS[k]!;
          return m.top === geo.marginPx.top && m.right === geo.marginPx.right &&
            m.bottom === geo.marginPx.bottom && m.left === geo.marginPx.left;
        }) ?? "Normal";
      colsSel.value = geo.columns?.count === 2 ? "Two" : geo.columns?.count === 3 ? "Three" : "One";
      startInput.value = geo.pageNumberStart === null ? "" : String(geo.pageNumberStart);
      const sec = editor.getDocument().section;
      firstCheck.checked = sec.headerFirst !== undefined || sec.footerFirst !== undefined;
      evenCheck.checked = sec.headerEven !== undefined || sec.footerEven !== undefined;
      panel.style.display = "flex";
    },
  };
})();

// ---- find & replace bar (Ctrl+F) -------------------------------------------
{
  const bar = document.createElement("div");
  bar.style.cssText =
    "position:fixed;top:46px;right:24px;display:none;gap:4px;align-items:center;" +
    "background:#fff;border:1px solid #dadce0;border-radius:8px;padding:6px 8px;" +
    "box-shadow:0 2px 8px rgba(0,0,0,.18);z-index:10;font-size:13px;";
  const mkInput = (placeholder: string, width: number): HTMLInputElement => {
    const i = document.createElement("input");
    i.placeholder = placeholder;
    i.style.cssText = `width:${width}px;height:24px;border:1px solid #dadce0;border-radius:4px;padding:0 6px;font-size:13px;`;
    bar.appendChild(i);
    return i;
  };
  const mkBtn = (label: string, title: string, onClick: () => void): HTMLButtonElement => {
    const b = document.createElement("button");
    b.textContent = label;
    b.title = title;
    b.style.cssText = "height:24px;border:none;background:transparent;cursor:pointer;font-size:13px;color:#3c4043;";
    b.addEventListener("click", onClick);
    bar.appendChild(b);
    return b;
  };
  const findInput = mkInput("Find", 140);
  const counter = document.createElement("span");
  counter.style.cssText = "color:#80868b;min-width:40px;text-align:center;";
  counter.textContent = "";
  bar.appendChild(counter);
  const update = (s: { index: number; total: number }): void => {
    counter.textContent = s.total > 0 ? `${s.index}/${s.total}` : findInput.value ? "0/0" : "";
  };
  // Match-case / whole-word toggles (editor.search already accepts these).
  let matchCase = false;
  let wholeWord = false;
  const reSearch = (): void => {
    update(findInput.value ? editor.search(findInput.value, { matchCase, wholeWord }) : { index: 0, total: 0 });
  };
  const optBtn = (label: string, title: string, get: () => boolean, set: (v: boolean) => void): void => {
    const b = document.createElement("button");
    b.textContent = label;
    b.title = title;
    b.style.cssText = "height:24px;min-width:26px;border:1px solid transparent;border-radius:4px;background:transparent;cursor:pointer;font-size:13px;font-weight:600;";
    const paint = (): void => {
      const on = get();
      b.style.background = on ? "#cfe3fb" : "transparent";
      b.style.borderColor = on ? "#b3d3f5" : "transparent";
      b.style.color = on ? "#0b57d0" : "#3c4043";
    };
    b.addEventListener("click", () => {
      set(!get());
      paint();
      reSearch();
      findInput.focus();
    });
    paint();
    bar.appendChild(b);
  };
  optBtn("Aa", "Match case", () => matchCase, (v) => (matchCase = v));
  optBtn("⌈W⌋", "Match whole word only", () => wholeWord, (v) => (wholeWord = v));
  mkBtn("‹", "Previous (Shift+Enter)", () => update(editor.searchNav(-1)));
  mkBtn("›", "Next (Enter)", () => update(editor.searchNav(1)));
  const replaceInput = mkInput("Replace", 120);
  mkBtn("Replace", "Replace current", () => update(editor.searchReplaceCurrent(replaceInput.value)));
  mkBtn("All", "Replace all", () => {
    const n = editor.searchReplaceAll(replaceInput.value);
    update(editor.search(findInput.value, { matchCase, wholeWord }));
    counter.textContent += ` (${n} replaced)`;
  });
  const close = (): void => {
    bar.style.display = "none";
    editor.searchClear();
    editor.focus();
  };
  mkBtn("×", "Close (Esc)", close);
  document.body.appendChild(bar);

  openFind = (): void => {
    bar.style.display = "flex";
    findInput.select();
    findInput.focus();
    reSearch();
  };

  findInput.addEventListener("input", reSearch);
  bar.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter") {
      update(editor.searchNav(ev.shiftKey ? -1 : 1));
      ev.preventDefault();
    } else if (ev.key === "Escape") {
      close();
      ev.preventDefault();
    }
    ev.stopPropagation(); // typing in the bar never reaches the editor keymap
  });
  window.addEventListener("keydown", (ev) => {
    if ((ev.ctrlKey || ev.metaKey) && ev.key.toLowerCase() === "f") {
      ev.preventDefault();
      openFind();
    }
  });
}

// Dev hook for in-browser verification (break-rule scans, perf probes).
declare global {
  interface Window {
    __cw?: unknown;
  }
}
window.__cw = { doc, tree, engine, editor, createLayoutEngine, sampleDoc, stressDoc };
