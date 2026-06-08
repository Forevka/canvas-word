// @cw/backend — collaboration server (HTTP + WebSocket) over a Postgres-backed
// change log. Filled in during Phase 3 (ChangeStore + server) and Phase 4 (OT
// ordering/broadcast). For now it just proves the shared-core wiring compiles.

import { applyOp } from "@cw/shared";

// Placeholder so the package typechecks before the server lands. Referencing a
// shared export keeps the @cw/shared resolution path exercised from the backend.
export const _wiringCheck = typeof applyOp === "function";
