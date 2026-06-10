import { defineConfig } from "vite";

// Deliberately minimal — embedding @forevka/wordcanvas needs no node polyfills,
// no aliases, no worker plumbing. The published bundle is self-contained (deps +
// the document core are inlined) and ships its own pre-built workers.
//
// Key point: the library is a *pre-built, code-split* ES bundle that loads its
// editor app + export/import workers via `new URL(..., import.meta.url)`. A
// downstream bundler must NOT re-process it, so we keep it EXTERNAL:
//   - dev   → `optimizeDeps.exclude` serves the bundle as-is from node_modules.
//   - build → `rollupOptions.external` leaves the bare import in the output; the
//             import map in index.html resolves it to the copied dist-lib at
//             runtime (see the `build` script's copy step + README).
//
// `base: "./"` keeps built asset paths relative, so the same build works at the
// dev root (`/`) and deployed under a subpath (e.g. Caddy's `/live`).
export default defineConfig({
  base: "./",
  server: { port: 5180 },
  optimizeDeps: { exclude: ["@forevka/wordcanvas"] },
  build: {
    // es2022 for the top-level `await promptIdentity()` in main.ts.
    target: "es2022",
    rollupOptions: { external: ["@forevka/wordcanvas"] },
  },
});
