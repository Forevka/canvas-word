// Demonstrates that several WordCanvas editors coexist on ONE page, each fully
// independent (own document, undo history, toolbar, and selection). The library's
// chrome is class-scoped under a per-instance root and each instance owns its
// off-container floats, so mounting and destroying instances is leak-free.
//
// Consumes the published @forevka/wordcanvas package exactly the way a third-party
// integrator would: import the class, mount it into a sized container, destroy it
// when the container goes away. Offline mode — no backend.

import { WordCanvas } from "@forevka/wordcanvas";

const editorsEl = document.getElementById("editors")!;
const countEl = document.getElementById("count")!;

// Track instances so "Remove last" can tear one down (and prove destroy() cleans
// up its chrome + body-level floating panels).
const instances: { pane: HTMLElement; editor: WordCanvas }[] = [];

const updateCount = (): void => {
  countEl.textContent = `${instances.length} editor${instances.length === 1 ? "" : "s"}`;
};

const addEditor = (): void => {
  const pane = document.createElement("div");
  pane.className = "pane";
  editorsEl.appendChild(pane);
  // Each instance is a separate, self-contained editor mounted into its own box.
  const editor = new WordCanvas({ container: pane });
  instances.push({ pane, editor });
  updateCount();
};

const removeLast = (): void => {
  const last = instances.pop();
  if (!last) return;
  last.editor.destroy(); // removes the editor chrome + its body-level floats
  last.pane.remove();
  updateCount();
};

document.getElementById("add")!.addEventListener("click", addEditor);
document.getElementById("remove")!.addEventListener("click", removeLast);

// Start with two side-by-side editors.
addEditor();
addEditor();

// Expose for in-browser verification.
(window as unknown as { __instances?: unknown }).__instances = instances;
