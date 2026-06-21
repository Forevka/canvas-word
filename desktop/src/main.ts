// WordCanvas desktop entry. Mounts the editor fully OFFLINE (no backendUrl → no
// network, no sync, no share). All document actions live in the editor's ribbon:
//   - Open .docx  → the ribbon's File ▸ Open uses the OS file picker and shows
//                   the editor's built-in "Opening document…" progress overlay.
//   - Export DOCX/PDF → the ribbon's Export buttons render (with a busy overlay)
//                   then hand the file to onSave below, which writes it through a
//                   native Save dialog instead of an unreliable webview download.
// So there's no separate native menu — the ribbon already covers everything.
import { WordCanvas } from "@forevka/wordcanvas";
import { isTauri, saveBytesViaDialog, type SaveFormat } from "./native";

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
  // Expose to AI agents over WebMCP (navigator.modelContext). Drop if unwanted.
  agentTools: true,
  onLoadProgress: ({ phase, percent }) => {
    bar.style.width = `${Math.round(percent * 100)}%`;
    label.textContent = PHASE_LABEL[phase] ?? "";
    if (phase === "ready") loader.classList.add("done");
  },
  // In Tauri, route the ribbon's Export through a native Save dialog. In a plain
  // browser this hook is omitted, so the default download applies.
  ...(tauri
    ? { onSave: async ({ bytes, format }) => saveBytesViaDialog(bytes, format as SaveFormat) }
    : {}),
});

// Expose for quick poking / Playwright.
(window as unknown as { __wc: WordCanvas }).__wc = editor;

editor.on("ready", () => console.log(`WordCanvas ready (${tauri ? "desktop" : "browser"})`));
