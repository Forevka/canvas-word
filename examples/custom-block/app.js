// Custom block-type demo. `registerBlockType` adds a NEW document block kind that
// draws itself on the page canvas — here a small bar chart. A custom block is
// atomic (measures to a box, paginates like an image), non-editable (no caret),
// and holds only JSON `data`, so it serializes for free. It has no native OOXML,
// so `.docx` export degrades it to a placeholder paragraph + a warning.
import { WordCanvas, registerBlockType } from "@forevka/wordcanvas";

// --- register the block type (once, before/around mount) -----------------------
registerBlockType({
  type: "barchart",
  // Atomic: return the block's height for the available content width.
  measure: (_data, _ctx) => ({ height: 168 }),
  // Paint: the canvas context is translated to the block's top-left, so draw in
  // local [0,width] × [0,height] document px. (Never measure text/layout here.)
  paint: (ctx, { width, height }, data) => {
    const { title = "Chart", values = [] } = data ?? {};
    const padT = 30, padB = 22, padX = 16;
    const plotW = width - padX * 2;
    const plotH = height - padT - padB;
    // card background
    ctx.fillStyle = "#f8fafc";
    ctx.strokeStyle = "#e2e8f0";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.roundRect(0.5, 0.5, width - 1, height - 1, 8);
    ctx.fill();
    ctx.stroke();
    // title
    ctx.fillStyle = "#0f172a";
    ctx.font = "600 14px system-ui, sans-serif";
    ctx.fillText(title, padX, 20);
    // Loading state — drawn until the (async) data arrives. The block paints this
    // synchronously; the FETCH happens outside, then `handle.setCustomBlockData`
    // swaps in the ready data and the editor re-renders (see the async button).
    if (data?.state === "loading") {
      ctx.fillStyle = "#94a3b8";
      ctx.font = "13px system-ui, sans-serif";
      ctx.textAlign = "center";
      ctx.fillText("Loading…", width / 2, height / 2 + 6);
      ctx.textAlign = "start";
      return;
    }
    // bars
    const max = Math.max(1, ...values.map((v) => v.value));
    const n = values.length || 1;
    const gap = 10;
    const barW = Math.max(4, (plotW - gap * (n - 1)) / n);
    values.forEach((v, i) => {
      const h = (v.value / max) * (plotH - 16);
      const x = padX + i * (barW + gap);
      const y = padT + (plotH - h);
      ctx.fillStyle = "#2563eb";
      ctx.beginPath();
      ctx.roundRect(x, y, barW, h, 3);
      ctx.fill();
      ctx.fillStyle = "#475569";
      ctx.font = "11px system-ui, sans-serif";
      ctx.textAlign = "center";
      ctx.fillText(String(v.label ?? ""), x + barW / 2, height - 8);
      ctx.textAlign = "start";
    });
  },
  // OPTIONAL: emit OOXML for .docx export. Omitted here on purpose, to demonstrate
  // the default lossy behavior (a placeholder paragraph + an export warning).
});

// A helper to build a fresh custom block for the model.
let seq = 0;
const chartBlock = (title, values) => ({
  kind: "custom",
  id: `chart-${seq++}-${Math.round(performance.now())}`,
  revision: 0,
  customType: "barchart",
  data: { title, values },
});

// A chart block that starts in a LOADING state (its `data` has no values yet).
const loadingChartBlock = (title) => ({
  kind: "custom",
  id: `chart-${seq++}-${Math.round(performance.now())}`,
  revision: 0,
  customType: "barchart",
  data: { title, state: "loading" },
});

const SAMPLE = [
  { label: "Q1", value: 42 },
  { label: "Q2", value: 68 },
  { label: "Q3", value: 55 },
  { label: "Q4", value: 91 },
];

// --- first-load progress bar ---------------------------------------------------
const loader = document.getElementById("loader");
const bar = loader.querySelector(".bar");
const label = loader.querySelector(".label");
const PHASE_LABEL = { bundle: "Loading editor…", fonts: "Loading fonts…", ready: "Ready" };

const editor = new WordCanvas({
  container: document.getElementById("editor"),
  onLoadProgress: ({ phase, percent }) => {
    bar.style.width = `${Math.round(percent * 100)}%`;
    label.textContent = PHASE_LABEL[phase];
    if (phase === "ready") loader.classList.add("done");
  },
});

// Insert a block into the body after the block that holds the caret (or at the end).
const insertAfterCaret = (doc, block) => {
  const sel = editor.__handle?.getSelection?.();
  const caretId = sel?.focus.blockId;
  const idx = caretId ? doc.blocks.findIndex((b) => b.id === caretId) : -1;
  const at = idx >= 0 ? idx + 1 : doc.blocks.length;
  doc.blocks.splice(at, 0, block);
  return doc;
};

editor.whenReady().then((handle) => {
  editor.__handle = handle;
  window.__handle = handle;

  // Insert one chart near the top on load (right after the first paragraph).
  const doc = handle.getDocument();
  const firstPara = doc.blocks.findIndex((b) => b.kind === "paragraph");
  doc.blocks.splice(firstPara >= 0 ? firstPara + 1 : 0, 0, chartBlock("Quarterly revenue", SAMPLE));
  handle.setDocument(doc);

  document.getElementById("add").onclick = () => {
    const d = handle.getDocument();
    insertAfterCaret(d, chartBlock("Chart " + (seq + 1), SAMPLE.map((v) => ({ ...v, value: Math.round(20 + Math.random() * 80) }))));
    handle.setDocument(d);
  };

  // Async pattern: insert a LOADING block, fetch data in your own async code, then
  // `handle.setCustomBlockData(id, …)` swaps it in and re-renders — an undoable
  // edit that bumps only this block (no full-document rebuild). paint()/measure()
  // stay synchronous; the async lives entirely in YOUR code, out here.
  document.getElementById("async").onclick = async () => {
    const block = loadingChartBlock("Revenue (live)");
    const d = handle.getDocument();
    insertAfterCaret(d, block);
    handle.setDocument(d); // shows "Loading…" immediately

    // …fetch from your backend (simulated here with a 1.2s delay) …
    const values = await new Promise((res) =>
      setTimeout(() => res(SAMPLE.map((v) => ({ ...v, value: Math.round(20 + Math.random() * 80) }))), 1200),
    );

    handle.setCustomBlockData(block.id, { title: "Revenue (live)", state: "ready", values });
    console.log("Data arrived → setCustomBlockData re-rendered the block (Ctrl/Cmd+Z to undo just this update)");
  };

  document.getElementById("export").onclick = async () => {
    const blob = await handle.exportDocx();
    // The block has no toOOXML, so export drops it to a placeholder paragraph and
    // records a `custom-block-dropped` warning (surfaced via the toolbar / onSave).
    const url = URL.createObjectURL(blob);
    const a = Object.assign(document.createElement("a"), { href: url, download: "with-custom-block.docx" });
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    console.log("Exported — the custom block was replaced by a placeholder paragraph (lossy).");
  };
});

editor.on("ready", () => console.log("WordCanvas ready — a custom 'barchart' block is registered"));
window.__wc = editor;
