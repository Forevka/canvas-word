// Dev harness for `npm run dev` — mounts WordCanvas into the page. Reads
// ?backend / ?collab so the local dev page can exercise online mode; production
// embedders instead do `new WordCanvas({ container, backendUrl })` themselves.

import { showIdentityPopup } from "./app/identityPopup";
import { WordCanvas, type LoadProgress } from "./wordcanvas";

// First-load loading bar. On a cold load the editor JS chunk and the bundled
// fonts (~9 MB) stream before the editor is interactive; `onLoadProgress` drives
// the #loader overlay (see index.html) so the page shows motion, not a blank box.
const loader = document.getElementById("loader");
const bar = loader?.querySelector<HTMLElement>(".bar");
const label = loader?.querySelector<HTMLElement>(".label");
const PHASE_LABEL: Record<LoadProgress["phase"], string> = {
  bundle: "Loading editor…",
  fonts: "Loading fonts…",
  ready: "Ready",
};
const onLoadProgress = ({ phase, percent }: LoadProgress): void => {
  if (bar) bar.style.width = `${Math.round(percent * 100)}%`;
  if (label) label.textContent = PHASE_LABEL[phase];
  if (phase === "ready") loader?.classList.add("done"); // CSS fades it out
};

const params = new URLSearchParams(location.search);
// `?backend=` wins; otherwise fall back to VITE_BACKEND (set by `npm run
// dev:online`), so online mode works without appending the query param.
const backend =
  params.get("backend") ?? (import.meta.env as Record<string, string | undefined>).VITE_BACKEND ?? null;
// `?doc` is the canonical param (e.g. the upload redirect target); `?collab` is
// the older alias kept working.
const collab = params.get("doc") ?? params.get("collab");

// `?view=<url-encoded JSON>` lets the dev page exercise the embedder `view`
// options (panels, rulers, grid, ribbon, zoom). Ignored if absent or malformed.
let view: import("./wordcanvas").WordCanvasViewOptions | undefined;
const viewParam = params.get("view");
if (viewParam) {
  try {
    view = JSON.parse(viewParam) as import("./wordcanvas").WordCanvasViewOptions;
  } catch {
    console.warn("[dev] ignoring malformed ?view= JSON");
  }
}

// `?onSave=1` wires a demo host save hook (records the last save on
// window.__saved) so the dev page can exercise the "host pipeline" save-state
// path; production embedders pass their own `onSave`. `?title=` sets an
// explicit document title.
const useOnSave = params.get("onSave") === "1";
const titleParam = params.get("title");

// Online (a backend is configured): ask who you are so edits/carets are
// attributed. Offline: no identity needed.
const user = backend ? await showIdentityPopup() : undefined;

// Demo roster so @-mentions in comments are exercisable in the dev harness;
// production embedders pass their own `knownUsers`.
const knownUsers = [
  { id: "u-ada", firstName: "Ada", lastName: "Lovelace" },
  { id: "u-alan", firstName: "Alan", lastName: "Turing" },
  { id: "u-grace", firstName: "Grace", lastName: "Hopper" },
  { id: "u-linus", firstName: "Linus", lastName: "Torvalds" },
];

const editor = new WordCanvas({
  container: document.body,
  ...(backend ? { backendUrl: backend } : {}),
  ...(collab ? { collabId: collab } : {}),
  ...(user ? { user } : {}),
  ...(view ? { view } : {}),
  ...(titleParam ? { documentTitle: titleParam } : {}),
  ...(useOnSave
    ? {
        onSave: (e: { format: string; bytes: Uint8Array }) => {
          (window as unknown as { __saved?: unknown }).__saved = { format: e.format, size: e.bytes.length, at: Date.now() };
        },
      }
    : {}),
  // `?devMode=true` reveals the Developer tab + Document-tree inspector.
  ...(params.get("devMode") === "true" ? { develop: true } : {}),
  // `?chrome=minimal` switches to the quiet command bar (Move 1); default ribbon.
  ...(params.get("chrome") === "minimal" ? { chrome: "minimal" as const } : params.get("chrome") === "ribbon" ? { chrome: "ribbon" as const } : {}),
  knownUsers,
  onLoadProgress,
});

// Expose for in-browser verification (Playwright).
(window as unknown as { __wc?: unknown }).__wc = editor;
