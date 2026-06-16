// One-command local "online" dev: Postgres (docker) + backend (:8787) + editor
// (:5173, VITE_BACKEND preset → online mode: Share, live collaboration, track
// changes + comments sync) + admin dashboard (:5174, for login + minting API
// tokens / managing webhooks). Dependency-free (just node + docker + npm).
//
//   npm run dev:online   → editor http://localhost:5173/ · dashboard http://localhost:5174/
//
import { spawn, spawnSync } from "node:child_process";

const BACKEND = process.env.BACKEND_URL ?? "http://localhost:8787";
const EDITOR = process.env.EDITOR_URL ?? "http://localhost:5173";
const sh = (cmd, opts = {}) => spawn(cmd, { stdio: "inherit", shell: true, ...opts });

console.log("[dev:online] starting Postgres (docker compose up -d)…");
const db = spawnSync("docker compose up -d", { stdio: "inherit", shell: true });
if (db.status !== 0) {
  console.error("[dev:online] could not start the database — is Docker Desktop running?");
  process.exit(1);
}

console.log(`[dev:online] backend   → ${BACKEND}`);
console.log(`[dev:online] editor    → ${EDITOR}/ (VITE_BACKEND preset)`);
console.log("[dev:online] dashboard → http://localhost:5174/ (admin login + API tokens)");
const backend = sh("npm run dev --workspace @cw/backend");
const frontend = sh("npm run dev --workspace @forevka/wordcanvas", {
  env: { ...process.env, VITE_BACKEND: BACKEND },
});
const dashboard = sh("npm run dev --workspace @cw/dashboard", {
  env: { ...process.env, VITE_BACKEND_URL: BACKEND, VITE_EDITOR_URL: EDITOR },
});

const procs = { backend, frontend, dashboard };
let exiting = false;
const stop = (code = 0) => {
  if (exiting) return;
  exiting = true;
  for (const p of Object.values(procs)) p.kill();
  process.exit(code);
};
process.on("SIGINT", () => stop(0));
process.on("SIGTERM", () => stop(0));
for (const [name, p] of Object.entries(procs)) {
  p.on("exit", (c) => { console.log(`[dev:online] ${name} exited (${c})`); stop(c ?? 0); });
}
