// Command + keymap registry demo. The `commands` option registers named commands
// and (optionally) binds keyboard shortcuts to them. Each command's `run` receives
// the same CommandContext a custom ribbon button gets (the editor handle + macro
// helpers: getSelection, insertText, emit, registerCleanup), so the SAME macro
// works from a keystroke, a ribbon button, or handle.runCommand(id).
//
// Built-in editing chords (Ctrl+B/I/U/Z/Y, Ctrl+Enter) always win — the registry
// skips any binding the built-in keymap already handled.
import { WordCanvas } from "@forevka/wordcanvas";

// --- a tiny on-screen event log so the commands show a visible effect ----------
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

// --- macros used by the commands (pure functions of the document) --------------

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

/** Count words + characters across the document. */
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

// --- mount ---------------------------------------------------------------------
const editor = new WordCanvas({
  container: document.getElementById("editor"),
  onLoadProgress: ({ phase, percent }) => {
    bar.style.width = `${Math.round(percent * 100)}%`;
    label.textContent = PHASE_LABEL[phase];
    if (phase === "ready") loader.classList.add("done");
  },

  // The whole feature: an array of commands, each with an optional keybinding.
  commands: [
    {
      id: "demo.uppercaseDoc",
      label: "Uppercase document",
      keybinding: "Alt+Shift+U",
      run: (ctx) => {
        // A macro: read the doc, transform a copy, write it back. setDocument
        // starts a fresh document (drops undo history) — fine for a batch macro.
        ctx.setDocument(uppercaseDoc(structuredClone(ctx.getDocument())));
        ctx.emit("uppercased");
      },
    },
    {
      id: "demo.insertDate",
      label: "Insert today's date",
      keybinding: "Alt+Shift+D",
      run: (ctx) => {
        ctx.insertText(new Date().toLocaleDateString());
        ctx.emit("inserted-date");
      },
    },
    {
      id: "demo.insertMarker",
      label: "Insert review marker",
      // A command can declare more than one shortcut.
      keybinding: ["Alt+Shift+M", "Alt+Shift+R"],
      run: (ctx) => {
        ctx.insertText(" [reviewed] ");
        ctx.emit("inserted-marker");
      },
    },
    {
      id: "demo.wordCount",
      label: "Word count",
      keybinding: "Mod+Shift+L", // Mod = Ctrl on win/linux, Cmd on mac
      run: (ctx) => {
        const { words, chars } = docStats(ctx.getDocument());
        ctx.emit("word-count", { words, chars });
      },
    },
  ],
});

// Commands surface their work through the standard `custom` event channel (emit).
editor.on("custom", ({ name, payload }) => {
  log(payload ? `${name} ${JSON.stringify(payload)}` : name);
});

editor.on("ready", () => console.log("WordCanvas ready — try Alt+Shift+U / D / M, or Ctrl/Cmd+Shift+L"));

// Commands are also invokable programmatically via the EditorHandle.
editor.whenReady().then((handle) => {
  window.__handle = handle; // for quick in-console poking: __handle.runCommand("demo.wordCount")
  document.getElementById("run-date").addEventListener("click", () => {
    const ok = handle.runCommand("demo.insertDate");
    if (!ok) log("runCommand: no such command");
  });
});

// Expose for quick in-browser poking.
window.__wc = editor;
