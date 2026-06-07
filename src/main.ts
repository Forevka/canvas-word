import { createEditor } from "./index";
import { createLayoutEngine } from "./layout/engine";
import { sampleDoc } from "./model/sampleDoc";
import { stressDoc } from "./model/stressDoc";
import { importDocx, type ImportResult } from "./import/docx/importDocx";

// Fonts must be resolved before the first layout — pretext measures with the same
// font strings the paint layer draws with, so a late font swap would desync them.
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
let editor = createEditor(app, doc, { engine, onChange: () => syncToolbar() });

// Replace the open document (docx import): tear down and rebuild the editor —
// the layout engine is reused so its caches survive across documents.
const replaceDocument = (next: typeof doc): void => {
  editor.destroy();
  doc = next;
  editor = createEditor(app, doc, { engine, onChange: () => syncToolbar() });
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
  toggleHighlight,
  toggleVerticalAlign,
  setLinkCmd,
  insertSectionBreak,
  applyPageSetup,
  pageSetupAt,
  setBandVariantEnabled,
  insertTocCmd,
  insertFootnoteCmd,
} from "./editor/commands";
import { defaultStylesheet } from "./model/stylesheet";

const TOOLBAR_SVG =
  "data:image/svg+xml," +
  encodeURIComponent(
    `<svg xmlns="http://www.w3.org/2000/svg" width="280" height="100">` +
      `<rect width="280" height="100" rx="8" fill="#34a853"/>` +
      `<text x="140" y="56" font-family="Arial" font-size="18" fill="#fff" text-anchor="middle">inserted image</text>` +
      `</svg>`,
  );

const toolbar = document.getElementById("toolbar");
if (toolbar) {
  const btn = (label: string, title: string, onClick: () => void, style = ""): void => {
    const b = document.createElement("button");
    b.textContent = label;
    b.title = title;
    if (style) b.setAttribute("style", style);
    // mousedown must not steal focus from the IME proxy
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", onClick);
    toolbar.appendChild(b);
  };
  const sep = (): void => {
    const s = document.createElement("div");
    s.className = "sep";
    toolbar.appendChild(s);
  };
  const select = (title: string, width: number): HTMLSelectElement => {
    const s = document.createElement("select");
    s.title = title;
    s.style.cssText = `height:28px;border:1px solid #dadce0;border-radius:4px;background:#fff;font-size:13px;width:${width}px;`;
    s.addEventListener("mousedown", (e) => e.stopPropagation());
    toolbar.appendChild(s);
    return s;
  };
  const opt = (s: HTMLSelectElement, value: string, label: string): void => {
    const o = document.createElement("option");
    o.value = value;
    o.textContent = label;
    s.appendChild(o);
  };

  // ---- named styles gallery ----
  const styleSelect = select("Paragraph style", 110);
  const stylesheet = (): ReturnType<typeof defaultStylesheet> =>
    editor.getDocument().stylesheet ?? defaultStylesheet();
  const rebuildStyleOptions = (): void => {
    styleSelect.textContent = "";
    for (const s of stylesheet().styles) opt(styleSelect, s.id, s.name);
  };
  rebuildStyleOptions();
  styleSelect.addEventListener("change", () => {
    editor.dispatch(applyNamedStyle(styleSelect.value));
    editor.focus();
  });
  btn("✎", "Update current style to match selection", () => {
    editor.dispatch(updateStyleToSelection(styleSelect.value));
  });
  btn("✚", "New style from selection…", () => {
    const name = prompt("New style name:");
    if (name) {
      editor.dispatch(createStyleFromSelection(name));
      rebuildStyleOptions();
      syncToolbar();
    }
  });
  sep();

  // ---- font family / size / line spacing ----
  const fontSelect = select("Font family", 130);
  const FONTS = [
    "Georgia, serif",
    "Arial, sans-serif",
    "Times New Roman, serif",
    "Verdana, sans-serif",
    "Trebuchet MS, sans-serif",
    "Consolas, monospace",
    "Courier New, monospace",
  ];
  for (const f of FONTS) opt(fontSelect, f, f.split(",")[0]!);
  fontSelect.addEventListener("change", () => {
    editor.setCharStyle({ fontFamily: fontSelect.value });
    editor.focus();
  });

  const sizeInput = document.createElement("input");
  sizeInput.type = "number";
  sizeInput.min = "6";
  sizeInput.max = "96";
  sizeInput.title = "Font size (px)";
  sizeInput.style.cssText = "width:48px;height:26px;border:1px solid #dadce0;border-radius:4px;font-size:13px;padding:0 4px;";
  sizeInput.addEventListener("mousedown", (e) => e.stopPropagation());
  sizeInput.addEventListener("change", () => {
    const v = Number(sizeInput.value);
    if (Number.isFinite(v) && v >= 6 && v <= 96) {
      editor.setCharStyle({ fontSizePx: v });
      editor.focus();
    }
  });
  toolbar.appendChild(sizeInput);

  const spacingSelect = select("Line spacing", 58);
  for (const v of ["1", "1.15", "1.35", "1.5", "2"]) opt(spacingSelect, v, `${v}×`);
  spacingSelect.addEventListener("change", () => {
    editor.dispatch(setParaProps({ lineHeight: Number(spacingSelect.value) }));
    editor.focus();
  });
  sep();

  // toolbar controls mirror the caret formatting
  syncToolbar = (): void => {
    const f = editor.currentFormat();
    if (f.styleId && [...styleSelect.options].some((o) => o.value === f.styleId)) {
      styleSelect.value = f.styleId;
    }
    if (f.fontFamily) {
      const match = FONTS.find((x) => x.toLowerCase() === f.fontFamily!.toLowerCase());
      if (match) fontSelect.value = match;
    }
    if (f.fontSizePx !== null && document.activeElement !== sizeInput) {
      sizeInput.value = String(f.fontSizePx);
    }
    if (f.lineHeight !== null) {
      const v = String(f.lineHeight);
      if ([...spacingSelect.options].some((o) => o.value === v)) spacingSelect.value = v;
    }
  };

  btn("↶", "Undo (Ctrl+Z)", () => editor.undo());
  btn("↷", "Redo (Ctrl+Y)", () => editor.redo());
  sep();
  btn("B", "Bold (Ctrl+B)", () => editor.toggleStyle("bold"), "font-weight:700");
  btn("I", "Italic (Ctrl+I)", () => editor.toggleStyle("italic"), "font-style:italic");
  btn("U", "Underline (Ctrl+U)", () => editor.toggleStyle("underline"), "text-decoration:underline");
  btn("S", "Strikethrough", () => editor.toggleStyle("strikethrough"), "text-decoration:line-through");
  btn("🖍", "Highlight (yellow)", () => editor.dispatch(toggleHighlight()));
  btn("x²", "Superscript", () => editor.dispatch(toggleVerticalAlign("super")));
  btn("x₂", "Subscript", () => editor.dispatch(toggleVerticalAlign("sub")));
  btn("🔗", "Insert/remove hyperlink", () => {
    const url = prompt("Link URL (empty to remove):");
    if (url !== null) editor.dispatch(setLinkCmd(url.trim() === "" ? null : url.trim()));
  });
  btn("🖌", "Format painter (double-click = sticky)", () => editor.armFormatPainter(false));
  sep();
  btn("•≡", "Bulleted list", () => editor.dispatch(toggleList("bullet")));
  btn("1≡", "Numbered list (Tab/Shift+Tab change level)", () => editor.dispatch(toggleList("decimal")));
  sep();
  // alignment routes to the selected image when one is selected
  btn("⫷", "Align left", () => editor.align("left"));
  btn("⫶", "Center", () => editor.align("center"));
  btn("⫸", "Align right", () => editor.align("right"));
  btn("☰", "Justify", () => editor.align("justify"));
  sep();
  btn("🖼", "Insert image", () => {
    editor.dispatch(insertImage(TOOLBAR_SVG, 280, 100)); // top-level caret
    editor.dispatch(insertImageInCell(TOOLBAR_SVG, 280, 100)); // table-cell caret
  });
  btn("⊞", "Insert 3×3 table", () => editor.dispatch(insertTable(3, 3)));
  sep();
  // table group — acts on the cell containing the caret (no-op outside tables)
  btn("⤒+", "Insert row above", () => editor.dispatch(insertTableRowCmd("above")));
  btn("⤓+", "Insert row below", () => editor.dispatch(insertTableRowCmd("below")));
  btn("⇤+", "Insert column left", () => editor.dispatch(insertTableColumnCmd("left")));
  btn("⇥+", "Insert column right", () => editor.dispatch(insertTableColumnCmd("right")));
  btn("─✕", "Delete row", () => editor.dispatch(deleteTableRowCmd()));
  btn("│✕", "Delete column", () => editor.dispatch(deleteTableColumnCmd()));
  btn("⊞✕", "Delete table", () => editor.dispatch(deleteTableCmd()));
  btn("⬚⬚", "Merge cells (select across cells in one row)", () => editor.dispatch(mergeCellsCmd()));
  btn("⬚|⬚", "Unmerge cell", () => editor.dispatch(unmergeCellCmd()));
  sep();
  // image wrap mode — acts on the selected image
  btn("◧", "Wrap text around image (square)", () => {
    const id = editor.getSelectedObject();
    if (id) editor.dispatch(setImageProps(id, { wrap: "square", align: "left" }));
  });
  btn("▣", "Image in line with text (block)", () => {
    const id = editor.getSelectedObject();
    if (id) editor.dispatch(setImageProps(id, { wrap: "block", align: "center" }));
  });
  sep();
  // sections & page setup
  btn("§⏎", "Insert section break (next page)", () => {
    editor.dispatch(insertSectionBreak());
    editor.focus();
  });
  btn("📐", "Page setup (applies to the caret's section)", () => {
    pageSetupPanel.toggle();
  });
  btn("☰§", "Insert / update table of contents (Ctrl+click an entry jumps to it)", () => {
    editor.dispatch(insertTocCmd());
    editor.focus();
  });
  btn("a¹", "Insert footnote", () => {
    editor.dispatch(insertFootnoteCmd());
    editor.focus();
  });
  sep();
  btn("📂", "Open .docx", () => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    input.addEventListener("change", () => {
      const file = input.files?.[0];
      if (file) void openDocxFile(file);
    });
    input.click();
  });
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
  mkBtn("↑", "Previous (Shift+Enter)", () => update(editor.searchNav(-1)));
  mkBtn("↓", "Next (Enter)", () => update(editor.searchNav(1)));
  const replaceInput = mkInput("Replace", 120);
  mkBtn("⇄", "Replace current", () => update(editor.searchReplaceCurrent(replaceInput.value)));
  mkBtn("⇄∀", "Replace all", () => {
    const n = editor.searchReplaceAll(replaceInput.value);
    update(editor.search(findInput.value));
    counter.textContent += ` (${n} replaced)`;
  });
  const close = (): void => {
    bar.style.display = "none";
    editor.searchClear();
    editor.focus();
  };
  mkBtn("✕", "Close (Esc)", close);
  document.body.appendChild(bar);

  findInput.addEventListener("input", () => update(editor.search(findInput.value)));
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
      bar.style.display = "flex";
      findInput.select();
      findInput.focus();
      if (findInput.value) update(editor.search(findInput.value));
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
