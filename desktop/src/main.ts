// WordCanvas desktop entry. Mounts the editor fully OFFLINE (no backendUrl → no
// network, no sync, no share) and — when running inside Tauri — routes file open
// and export through native OS dialogs.
import { WordCanvas } from "@forevka/wordcanvas";
import { isTauri, openDocxViaDialog, saveBytesViaDialog, type SaveFormat } from "./native";

// First-load loading bar. On a cold load the editor JS chunk and the bundled
// fonts (~9 MB) stream before the editor is interactive; onLoadProgress drives
// this bar. See the #loader markup in index.html.
const loader = document.getElementById("loader")!;
const bar = loader.querySelector<HTMLDivElement>(".bar")!;
const label = loader.querySelector<HTMLDivElement>(".label")!;
const PHASE_LABEL: Record<string, string> = { bundle: "Loading editor…", fonts: "Loading fonts…", ready: "Ready" };

const tauri = isTauri();

const editor = new WordCanvas({
  container: document.getElementById("editor")!,
  // Expose to AI agents over WebMCP (navigator.modelContext) — same as the
  // offline example. Drop if you don't want it.
  agentTools: true,
  onLoadProgress: ({ phase, percent }) => {
    bar.style.width = `${Math.round(percent * 100)}%`;
    label.textContent = PHASE_LABEL[phase] ?? "";
    if (phase === "ready") loader.classList.add("done");
  },
  // In Tauri, the editor's built-in Export (DOCX/PDF) buttons hand the produced
  // file here instead of triggering an unreliable webview anchor-download; we
  // write it through a native Save dialog. In a plain browser this hook is
  // omitted, so the default download behaviour applies.
  ...(tauri
    ? { onSave: async ({ bytes, format }) => saveBytesViaDialog(bytes, format as SaveFormat) }
    : {}),
});

// Expose for quick poking / Playwright.
(window as unknown as { __wc: WordCanvas }).__wc = editor;

editor.on("ready", () => console.log(`WordCanvas ready (${tauri ? "desktop" : "browser"})`));

if (tauri) {
  void editor.whenReady().then(setupNativeMenu);
  // Ctrl/Cmd+O works even if the app menu is unavailable.
  window.addEventListener("keydown", (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "o") {
      e.preventDefault();
      void openDocument();
    }
  });
}

async function openDocument(): Promise<void> {
  const buf = await openDocxViaDialog();
  if (buf) await editor.openDocx(buf);
}

async function saveAs(format: SaveFormat): Promise<void> {
  const blob = format === "pdf" ? await editor.exportPdf() : await editor.exportDocx();
  await saveBytesViaDialog(new Uint8Array(await blob.arrayBuffer()), format);
}

// Native application menu (File → Open / Save as DOCX / Save as PDF). Built from
// JS via the core menu API so no Rust menu code is needed. Uses permissions in
// core:default (capabilities/default.json).
async function setupNativeMenu(): Promise<void> {
  try {
    const { Menu } = await import("@tauri-apps/api/menu");
    const menu = await Menu.new({
      items: [
        {
          id: "file",
          text: "File",
          items: [
            { id: "open", text: "Open…", accelerator: "CmdOrCtrl+O", action: () => void openDocument() },
            { item: "Separator" },
            { id: "save_docx", text: "Save as DOCX…", accelerator: "CmdOrCtrl+S", action: () => void saveAs("docx") },
            { id: "save_pdf", text: "Save as PDF…", action: () => void saveAs("pdf") },
          ],
        },
      ],
    });
    await menu.setAsAppMenu();
  } catch (err) {
    // Menu is a nice-to-have; Ctrl+O and the in-editor Export buttons still work.
    console.warn("native menu unavailable", err);
  }
}
