// Entry for the dedicated "/live" page: instantiates WordCanvas in ONLINE mode
// with a hardcoded backend URL — the same way a real embedder would, instead of
// passing ?backend= on the dev page. Point BACKEND_URL at your API; here it
// targets the local docker-compose backend.
//
// In production an embedder would use the published package:
//   import { WordCanvas } from "@cw/frontend";
//   new WordCanvas({ container, backendUrl: "https://api.example.com" });

import { WordCanvas } from "./wordcanvas";

const BACKEND_URL = "http://localhost:8787";

// Still honor ?collab so a generated share link opens here and joins the session.
const collab = new URLSearchParams(location.search).get("collab");

const editor = new WordCanvas({
  container: document.body,
  backendUrl: BACKEND_URL,
  ...(collab ? { collabId: collab } : {}),
});

// Expose for in-browser verification.
(window as unknown as { __wc?: unknown }).__wc = editor;
