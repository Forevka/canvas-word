// Entry for the dedicated "/live" page: instantiates WordCanvas in ONLINE mode
// with a hardcoded backend URL — the same way a real embedder would, instead of
// passing ?backend= on the dev page. Point BACKEND_URL at your API; here it
// targets the local docker-compose backend.
//
// In production an embedder would use the published package:
//   import { WordCanvas } from "@cw/frontend";
//   new WordCanvas({ container, backendUrl: "https://api.example.com" });

import { showIdentityPopup } from "./app/identityPopup";
import { WordCanvas } from "./wordcanvas";

// Dev: hit the local docker-compose backend directly. Production (built): the
// frontend and backend share one origin behind Caddy, so derive it from the
// page — `https://doc-editor.example/` ⇒ API there, WS at `wss://…/ws`.
const BACKEND_URL = import.meta.env.DEV ? "http://localhost:8787" : location.origin;

// Still honor ?collab so a generated share link opens here and joins the session.
const collab = new URLSearchParams(location.search).get("collab");

// Online demo: ask who you are (POC identity) so carets/edits are attributed.
// A real embedder would pass `user` directly and skip this.
const user = await showIdentityPopup();

const editor = new WordCanvas({
  container: document.body,
  backendUrl: BACKEND_URL,
  user,
  ...(collab ? { collabId: collab } : {}),
});

// Expose for in-browser verification.
(window as unknown as { __wc?: unknown }).__wc = editor;
