// Entry point (tsx src/server.ts) — the only place that touches a real Postgres
// pool and binds a port. The app itself is assembled in app.ts; tests import
// createApp from there with their own store.
import { createApp } from "./app";
import { PORT } from "./config";
import { createPool } from "./db";
import { PgChangeStore } from "./store/ChangeStore";

const pool = createPool();
const { app } = await createApp(new PgChangeStore(pool));
await app.listen({ port: PORT, host: "0.0.0.0" });
console.log(`[cw-backend] listening on http://localhost:${PORT} (ws: /ws?doc=<id>)`);
