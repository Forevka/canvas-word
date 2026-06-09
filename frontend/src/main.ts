// Dev harness for `npm run dev` — mounts WordCanvas into the page. Reads
// ?backend / ?collab so the local dev page can exercise online mode; production
// embedders instead do `new WordCanvas({ container, backendUrl })` themselves.

import { WordCanvas } from "./wordcanvas";

const params = new URLSearchParams(location.search);
const backend = params.get("backend");
const collab = params.get("collab");

const editor = new WordCanvas({
  container: document.body,
  ...(backend ? { backendUrl: backend } : {}),
  ...(collab ? { collabId: collab } : {}),
});

// Expose for in-browser verification (Playwright).
(window as unknown as { __wc?: unknown }).__wc = editor;
