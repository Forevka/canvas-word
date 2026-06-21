// Post-build: copy the built @forevka/wordcanvas bundle next to the app as
// `dist/wordcanvas/`, so the production `dist/` is self-contained and the
// index.html import map (`./wordcanvas/wordcanvas.js`) resolves under Tauri's
// asset protocol. Run after `vite build`. Requires the library to be built
// first (`npm run build:lib --workspace @forevka/wordcanvas` from the repo root).

import { cp, rm, access } from "node:fs/promises";
import { fileURLToPath } from "node:url";

const src = fileURLToPath(new URL("../../frontend/dist-lib", import.meta.url));
const dst = fileURLToPath(new URL("../dist/wordcanvas", import.meta.url));

try {
  await access(src);
} catch {
  console.error(`Library bundle not found at ${src}\nRun \`npm run build:lib --workspace @forevka/wordcanvas\` first.`);
  process.exit(1);
}

await rm(dst, { recursive: true, force: true });
await cp(src, dst, { recursive: true });
console.log(`copied dist-lib -> ${dst}`);
