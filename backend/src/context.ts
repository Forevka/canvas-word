// Shared dependencies handed to every route module — the seam that keeps the HTTP
// handlers (presentation) decoupled from how data, realtime, and webhooks are wired
// together in the composition root (app.ts).
import type { Broadcaster } from "./realtime/Broadcaster";
import type { ChangeStore } from "./store/ChangeStore";
import type { WebhookDispatcher } from "./webhooks/dispatcher";

export interface AppContext {
  store: ChangeStore;
  bcast: Broadcaster;
  dispatcher: WebhookDispatcher;
}
