// Online + collaboration demo, migrated from the old frontend `/live` page. It
// consumes the published @forevka/wordcanvas package exactly the way a third-party
// integrator would: import the class, mount it into a sized container, and pass a
// backend URL plus a user identity.

import { WordCanvas, type LoadProgress } from "@forevka/wordcanvas";
import { promptIdentity } from "./identity";

// First-load loading bar. On a cold load the editor JS chunk and the bundled
// fonts (~9 MB) stream before the editor is interactive; `onLoadProgress` lets us
// show progress instead of a blank container. See the #loader markup in index.html.
const loader = document.getElementById("loader")!;
const bar = loader.querySelector<HTMLElement>(".bar")!;
const label = loader.querySelector<HTMLElement>(".label")!;
const PHASE_LABEL: Record<LoadProgress["phase"], string> = {
  bundle: "Loading editor…",
  fonts: "Loading fonts…",
  ready: "Ready",
};
const onLoadProgress = ({ phase, percent }: LoadProgress): void => {
  bar.style.width = `${Math.round(percent * 100)}%`;
  label.textContent = PHASE_LABEL[phase];
  if (phase === "ready") loader.classList.add("done"); // CSS fades it out
};

// Point this at your backend. In dev we hit the local docker-compose backend
// directly; a production deploy behind one origin (Caddy) can derive it from the
// page — `https://doc-editor.example/` ⇒ API there, WS at `wss://…/ws`.
const BACKEND_URL = import.meta.env.DEV ? "http://localhost:8787" : location.origin;

// Honor ?collab=<docId> so a generated share link opens here and joins the session.
const collab = new URLSearchParams(location.search).get("collab");

// The embedder owns identity — a real app already knows its user and would skip
// this, passing `user` directly. The demo asks so carets/edits are attributed.
const user = await promptIdentity();

const editor = new WordCanvas({
  container: document.getElementById("editor")!,
  backendUrl: BACKEND_URL,
  user,
  onLoadProgress,
  // Expose this (online/collab) document to AI agents over WebMCP — the
  // "connect an agent as reviewer/editor to a live document" workflow. Connect via
  // the WebMCP browser extension / Chrome DevTools MCP, then have the agent read,
  // comment, suggest, or edit. `true` = all tools.
  agentTools: true,
  // `docId` joins an existing collaboration session; here it comes from the share
  // link's ?collab param.
  ...(collab ? { docId: collab } : {}),
});

editor.on("ready", () => console.log("WordCanvas ready (online)"));
editor.on("shared", ({ url }) => console.log("share link:", url));
editor.on("presence", ({ participants }) => console.log("participants:", participants.length));

// Expose for in-browser verification.
(window as unknown as { __wc?: unknown }).__wc = editor;
