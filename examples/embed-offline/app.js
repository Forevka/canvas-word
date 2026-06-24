// Minimal OFFLINE embed. No `backendUrl`, so WordCanvas runs fully local — no
// network, no sync, no share. This is the smallest possible integration: one
// import, one container, one `new`.
import { WordCanvas } from "@forevka/wordcanvas";

// First-load loading bar. On a cold load the editor JS chunk and the bundled
// fonts (~9 MB) stream before the editor is interactive; `onLoadProgress` lets us
// show progress instead of a blank container. See the #loader markup in index.html.
const loader = document.getElementById("loader");
const bar = loader.querySelector(".bar");
const label = loader.querySelector(".label");
const PHASE_LABEL = { bundle: "Loading editor…", fonts: "Loading fonts…", ready: "Ready" };

// Optional embedder `view` config via ?view=<url-encoded JSON>. Lets a host page
// (e.g. the landing-page <iframe> embed) set the initial chrome state — hide the
// outline pane, start at 50% zoom, etc. Ignored if absent or malformed.
const params = new URLSearchParams(location.search);
let view;
const viewParam = params.get("view");
if (viewParam) {
  try {
    view = JSON.parse(viewParam);
  } catch {
    console.warn("ignoring malformed ?view= JSON");
  }
}

// `?devMode=true` reveals a dedicated **Developer** ribbon tab whose "Inspect
// document tree" button opens a floating, devtools-Elements-style panel over the
// parsed Document model: browse the tree (blocks → runs, tables → cells, bands,
// footnotes, side-tables), hover a node to highlight its region on the page, hover
// the page to reveal the matching node, click to jump to it, and read each node's
// raw JSON. Pure debugging aid — gated, so production embeds that omit it pay nothing.
const devMode = params.get("devMode") === "true";

const editor = new WordCanvas({
  container: document.getElementById("editor"),
  ...(view ? { view } : {}),
  ...(devMode ? { develop: true } : {}),
  // Drive the loading bar. `percent` is an overall 0..1 across the whole startup
  // (bundle download → font fetch → ready), so it maps straight onto a bar width.
  // `phase` is a coarse stage label; `loaded`/`total` are byte counts for the
  // measurable (fonts) phase if you want to show "3.2 / 9.1 MB".
  onLoadProgress: ({ phase, percent }) => {
    bar.style.width = `${Math.round(percent * 100)}%`;
    label.textContent = PHASE_LABEL[phase];
    if (phase === "ready") loader.classList.add("done"); // CSS fades it out
  },
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

// Expose for quick in-browser poking.
window.__wc = editor;
