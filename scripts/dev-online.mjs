// One-command local "online" dev: Postgres (docker) + backend (:8787) + frontend
// with VITE_BACKEND preset, so the editor mounts in online mode (Share button,
// live collaboration, track-changes + comments sync) without appending
// ?backend=… to the URL. Dependency-free (just node + docker + npm).
//
//   npm run dev:online      → open http://localhost:5173/
//
import { spawn, spawnSync } from "node:child_process";

const BACKEND = process.env.BACKEND_URL ?? "http://localhost:8787";
const sh = (cmd, opts = {}) => spawn(cmd, { stdio: "inherit", shell: true, ...opts });

console.log("[dev:online] starting Postgres (docker compose up -d)…");
const db = spawnSync("docker compose up -d", { stdio: "inherit", shell: true });
if (db.status !== 0) {
  console.error("[dev:online] could not start the database — is Docker Desktop running?");
  process.exit(1);
}

console.log(`[dev:online] backend → ${BACKEND}`);
console.log("[dev:online] frontend → http://localhost:5173/ (VITE_BACKEND preset)");
const backend = sh("npm run dev --workspace @cw/backend");
const frontend = sh("npm run dev --workspace @forevka/wordcanvas", {
  env: { ...process.env, VITE_BACKEND: BACKEND },
});

let exiting = false;
const stop = (code = 0) => {
  if (exiting) return;
  exiting = true;
  backend.kill();
  frontend.kill();
  process.exit(code);
};
process.on("SIGINT", () => stop(0));
process.on("SIGTERM", () => stop(0));
backend.on("exit", (c) => { console.log(`[dev:online] backend exited (${c})`); stop(c ?? 0); });
frontend.on("exit", (c) => { console.log(`[dev:online] frontend exited (${c})`); stop(c ?? 0); });
