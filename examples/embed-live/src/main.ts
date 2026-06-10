// Online + collaboration demo, migrated from the old frontend `/live` page. It
// consumes the published @forevka/wordcanvas package exactly the way a third-party
// integrator would: import the class, mount it into a sized container, and pass a
// backend URL plus a user identity.

import { WordCanvas } from "@forevka/wordcanvas";
import { promptIdentity } from "./identity";

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
  // `docId` joins an existing collaboration session; here it comes from the share
  // link's ?collab param.
  ...(collab ? { docId: collab } : {}),
});

editor.on("ready", () => console.log("WordCanvas ready (online)"));
editor.on("shared", ({ url }) => console.log("share link:", url));
editor.on("presence", ({ participants }) => console.log("participants:", participants.length));

// Expose for in-browser verification.
(window as unknown as { __wc?: unknown }).__wc = editor;
