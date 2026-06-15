// Minimal OFFLINE embed. No `backendUrl`, so WordCanvas runs fully local — no
// network, no sync, no share. This is the smallest possible integration: one
// import, one container, one `new`.
import { WordCanvas } from "@forevka/wordcanvas";

const editor = new WordCanvas({
  container: document.getElementById("editor"),
  // Expose the editor to AI agents over WebMCP (the standard navigator.modelContext
  // API). This is exactly the local-debugging workflow: open a document a user
  // reported as "rendering weirdly", then connect an agent (WebMCP browser
  // extension / Chrome DevTools MCP) and ask it to call `inspect_layout` to see the
  // page/line/fragment geometry, `replace_text`/`format_text` to fix content, or
  // `add_comment` to leave review notes. `true` = all tools; pass an object to
  // restrict, e.g. { capabilities: ["read"] } for read-only inspection.
  agentTools: true,
  // Custom fields: when a document contains a developer-defined field (e.g. a
  // paragraph whose w:instrText is ` MYCHART "sales-2026" `), right-clicking it
  // shows "Update Field (MYCHART)". This callback computes the field's result.
  //
  // In a real app you'd POST {name, instruction} to your backend and return the
  // **.docx it renders** (an ArrayBuffer) — that path is imported in full, so
  // images and rich content come through, e.g.:
  //   const res = await fetch(`/fields/${name}`, { method: "POST", body: instruction });
  //   return await res.arrayBuffer();
  // Here (offline, no backend) we return a simple OOXML fragment string instead,
  // which is also accepted for text-only results.
  resolveField: async ({ name, instruction }) => {
    const xmlEsc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    const text = `${name} resolved at ${new Date().toLocaleTimeString()} — instr:${instruction.trim()}`;
    return `<w:p><w:r><w:t xml:space="preserve">${xmlEsc(text)}</w:t></w:r></w:p>`;
  },
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
