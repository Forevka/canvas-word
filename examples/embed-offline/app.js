// Minimal OFFLINE embed. No `backendUrl`, so WordCanvas runs fully local — no
// network, no sync, no share. This is the smallest possible integration: one
// import, one container, one `new`.
import { WordCanvas } from "@forevka/wordcanvas";

const editor = new WordCanvas({
  container: document.getElementById("editor"),
});

editor.on("ready", () => {
  console.log("WordCanvas ready (offline)");
});

// Optional: let the user open a .docx from disk to see DOCX import (still local).
const picker = Object.assign(document.createElement("input"), { type: "file", accept: ".docx" });
picker.style.cssText = "position:fixed;z-index:10;top:8px;right:8px;";
picker.addEventListener("change", () => {
  const file = picker.files?.[0];
  if (file) void editor.openDocx(file);
});
document.body.appendChild(picker);

// Expose for quick in-browser poking.
window.__wc = editor;
