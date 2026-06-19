// Ribbon customization demo. The `customizeRibbon` option runs once at mount with
// a RibbonApi that addresses everything by id, so an embedder can reorder/remove
// the built-in tabs/groups/buttons and add their own — wiring custom macros,
// config popups, and informational popups. A custom button's `onClick` receives a
// RibbonActionContext: the editor handle PLUS macro helpers (getSelection,
// insertText, emit, registerCleanup).
import { WordCanvas, makeFloatingDialog } from "@forevka/wordcanvas";

// --- a tiny on-screen event log so the demo buttons show a visible effect ------
const logEl = document.getElementById("log");
const log = (msg) => {
  logEl.querySelector(".empty")?.remove();
  const line = document.createElement("div");
  line.textContent = `• ${msg}`;
  logEl.appendChild(line);
  logEl.scrollTop = logEl.scrollHeight;
};

// --- first-load progress bar (see index.html #loader) --------------------------
const loader = document.getElementById("loader");
const bar = loader.querySelector(".bar");
const label = loader.querySelector(".label");
const PHASE_LABEL = { bundle: "Loading editor…", fonts: "Loading fonts…", ready: "Ready" };

// --- helpers used by the custom buttons ----------------------------------------

/** Uppercase every run of text in the document (a whole-document macro). */
const uppercaseDoc = (doc) => {
  const walk = (blocks) => {
    for (const b of blocks) {
      if (b.kind === "paragraph") for (const r of b.runs) r.text = r.text.toUpperCase();
      else if (b.kind === "table") for (const row of b.rows) for (const cell of row.cells) walk(cell.blocks);
    }
  };
  walk(doc.blocks);
  return doc;
};

/** Count words + characters across the document (for the info popup). */
const docStats = (doc) => {
  let text = "";
  const walk = (blocks) => {
    for (const b of blocks) {
      if (b.kind === "paragraph") { for (const r of b.runs) text += r.text; text += "\n"; }
      else if (b.kind === "table") for (const row of b.rows) for (const cell of row.cells) walk(cell.blocks);
    }
  };
  walk(doc.blocks);
  const words = (text.match(/\S+/g) ?? []).length;
  return { words, chars: text.replace(/\s/g, "").length };
};

/** Open a small draggable, non-blocking popup (config / informational). Reuses the
 *  editor's own `makeFloatingDialog`, and ties the DOM + listeners to the editor's
 *  teardown via `ctx.registerCleanup` so nothing leaks when the editor is destroyed. */
const openDialog = (ctx, title, bodyNode) => {
  const backdrop = document.createElement("div");
  backdrop.style.cssText = "position:fixed;inset:0;z-index:9999;";
  const modal = document.createElement("div");
  modal.className = "demo-dialog";
  const header = document.createElement("header");
  header.className = "cw-drag-handle"; // makeFloatingDialog drags by this element
  const titleEl = document.createElement("span");
  titleEl.textContent = title;
  const close = document.createElement("button");
  close.textContent = "×";
  close.title = "Close";
  header.append(titleEl, close);
  const body = document.createElement("div");
  body.className = "body";
  body.append(bodyNode);
  modal.append(header, body);
  backdrop.append(modal);
  document.body.append(backdrop);

  const ac = new AbortController();
  makeFloatingDialog({ backdrop, modal, handle: header, signal: ac.signal, noDrag: "button" });
  const dispose = () => { backdrop.remove(); ac.abort(); };
  close.addEventListener("click", dispose);
  ctx.registerCleanup(backdrop); // removed if the editor is destroyed first
  ctx.registerCleanup(() => ac.abort());
  return dispose;
};

// --- mount ---------------------------------------------------------------------
const editor = new WordCanvas({
  container: document.getElementById("editor"),
  onLoadProgress: ({ phase, percent }) => {
    bar.style.width = `${Math.round(percent * 100)}%`;
    label.textContent = PHASE_LABEL[phase];
    if (phase === "ready") loader.classList.add("done");
  },

  customizeRibbon: (api) => {
    // 1) DISCOVER built-in ids. Ids are namespaced `tab.group.item`; rather than
    //    hard-coding, inspect them at runtime. (Open the console to see these.)
    console.log("[ribbon] tabs:", api.tabs());
    console.log("[ribbon] home groups:", api.groups("home"));
    console.log("[ribbon] home.font items:", api.items("home.font"));

    // 2) REORDER / REMOVE built-ins by id (no-op + warning if an id is unknown).
    api.moveItem("home.font.italic", { before: "home.font.bold" }); // Italic before Bold
    api.removeItem("home.clipboard.cut"); // drop the Cut button

    // 3) ADD a custom tab with two groups of buttons, placed after View.
    api.addTab({ id: "automate", label: "Automate", after: "view" });

    api.addGroup("automate", { id: "automate.macros", label: "Macros" });
    api.addButton("automate.macros", {
      id: "automate.macros.date",
      icon: "📅",
      tooltip: "Insert today's date at the caret",
      enabled: (fmt) => !fmt.imageSelected, // greyed out while an image is selected
      onClick: (ctx) => {
        ctx.insertText(new Date().toLocaleDateString());
        ctx.emit("inserted-date");
      },
    });
    api.addButton("automate.macros", {
      id: "automate.macros.upper",
      label: "UPPER",
      tooltip: "Uppercase the whole document",
      onClick: (ctx) => {
        // A macro: read the document, transform a copy, write it back. setDocument
        // starts a fresh document (drops undo history) — perfect for a batch macro.
        void ctx.setDocument(uppercaseDoc(structuredClone(ctx.getDocument())));
        ctx.emit("uppercased");
      },
    });

    api.addGroup("automate", { id: "automate.tools", label: "Tools" });
    api.addButton("automate.tools", {
      id: "automate.tools.count",
      icon: "🔢",
      tooltip: "Word count…",
      onClick: (ctx) => {
        const { words, chars } = docStats(ctx.getDocument());
        const p = document.createElement("div");
        p.textContent = `${words} words · ${chars} characters`;
        openDialog(ctx, "Word count", p);
        ctx.emit("counted", { words, chars });
      },
    });
    api.addButton("automate.tools", {
      id: "automate.tools.export",
      icon: "⬇️",
      tooltip: "Export .docx to your own pipeline",
      onClick: async (ctx) => {
        // Custom buttons can drive the full handle API too — here, the same export
        // an embedder would POST to their backend; we just download it.
        const blob = await ctx.exportDocx();
        const url = URL.createObjectURL(blob);
        const a = Object.assign(document.createElement("a"), { href: url, download: "document.docx" });
        a.click();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        ctx.emit("exported");
      },
    });

    // 4) ADD a button into a BUILT-IN tab/group (Home ▸ Editing).
    api.addButton("home.editing", {
      id: "automate.home.stamp",
      icon: "✓",
      tooltip: "Insert a review stamp",
      onClick: (ctx) => ctx.insertText(" [reviewed] "),
    });
  },
});

// Custom buttons surface events through the standard `custom` event channel.
editor.on("custom", ({ name, payload }) => {
  log(payload ? `${name} ${JSON.stringify(payload)}` : name);
});

editor.on("ready", () => console.log("WordCanvas ready — try the Automate tab"));

// Expose for quick in-browser poking.
window.__wc = editor;
