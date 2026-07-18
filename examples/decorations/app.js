// Decorations API demo. `handle.setDecorations(specs)` draws custom paint-only
// overlays at DOCUMENT coordinates — highlight a range, underline/box it, or
// badge a position. You describe them in document space (a DocSelection range or
// a DocPosition); the editor resolves them against the live layout and repaints,
// re-resolving after every edit so they stay anchored to the text (never
// measuring or re-breaking lines — same feed as the track-changes overlays).
import { WordCanvas } from "@forevka/wordcanvas";

// --- first-load progress bar (see index.html #loader) --------------------------
const loader = document.getElementById("loader");
const bar = loader.querySelector(".bar");
const label = loader.querySelector(".label");
const PHASE_LABEL = { bundle: "Loading editor…", fonts: "Loading fonts…", ready: "Ready" };

// We keep the current decoration set locally and replace the whole set on each
// change — `setDecorations` is declarative (like React state), not incremental.
let decos = [];
let handle = null;
const apply = () => handle?.setDecorations(decos);

// --- mount ---------------------------------------------------------------------
const editor = new WordCanvas({
  container: document.getElementById("editor"),
  onLoadProgress: ({ phase, percent }) => {
    bar.style.width = `${Math.round(percent * 100)}%`;
    label.textContent = PHASE_LABEL[phase];
    if (phase === "ready") loader.classList.add("done");
  },
});

// A range decoration needs a DocSelection ({ anchor, focus }); the current
// selection is exactly that shape, so we reuse it directly.
const selectionRange = () => {
  const sel = handle.getSelection();
  if (!sel || (sel.anchor.blockId === sel.focus.blockId && sel.anchor.offset === sel.focus.offset)) return null;
  return { anchor: sel.anchor, focus: sel.focus };
};

const addRangeDeco = (type, color, extra = {}) => {
  const range = selectionRange();
  if (!range) return alert("Select some text first.");
  decos.push({ type, range, color, ...extra });
  apply();
};

editor.whenReady().then((h) => {
  handle = h;
  window.__handle = h; // for console poking

  // --- an initial decoration set drawn on load -------------------------------
  // Highlight the first ~14 chars of the first paragraph, and badge its start —
  // built straight from the document model (no selection needed).
  const doc = h.getDocument();
  const firstPara = doc.blocks.find((b) => b.kind === "paragraph" && b.runs.some((r) => r.text.trim()));
  if (firstPara) {
    const textLen = firstPara.runs.reduce((n, r) => n + r.text.length, 0);
    decos.push({
      type: "highlight",
      color: "#ffe082",
      range: { anchor: { blockId: firstPara.id, offset: 0 }, focus: { blockId: firstPara.id, offset: Math.min(14, textLen) } },
    });
    decos.push({ type: "badge", at: { blockId: firstPara.id, offset: 0 }, color: "#1e88e5", label: "1" });
    apply();
  }

  // --- panel wiring ----------------------------------------------------------
  document.getElementById("hl").onclick = () => addRangeDeco("highlight", "#a5d6a7", { opacity: 0.5 });
  document.getElementById("ul").onclick = () => addRangeDeco("underline", "#e53935", { thickness: 2 });
  document.getElementById("box").onclick = () => addRangeDeco("box", "#1e88e5", { thickness: 1.5 });
  document.getElementById("badge").onclick = () => {
    const sel = handle.getSelection();
    if (!sel) return alert("Place the caret somewhere first.");
    decos.push({ type: "badge", at: sel.focus, color: "#8e24aa", label: String(decos.filter((d) => d.type === "badge").length + 1) });
    apply();
  };
  document.getElementById("clear").onclick = () => {
    decos = [];
    handle.clearDecorations();
  };
});

editor.on("ready", () => console.log("WordCanvas ready — select text and add a decoration"));
window.__wc = editor;
