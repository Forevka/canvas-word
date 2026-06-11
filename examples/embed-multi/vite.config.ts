import { defineConfig } from "vite";

// Mirrors examples/embed-live: the published @forevka/wordcanvas bundle is a
// pre-built, code-split ES library that loads its editor app + import/export
// workers via `new URL(..., import.meta.url)`. A downstream bundler must NOT
// re-process it, so it stays EXTERNAL:
//   - dev   → `optimizeDeps.exclude` serves the bundle as-is from node_modules.
//   - build → `rollupOptions.external` leaves the bare import in the output; the
//             import map in index.html resolves it to the copied dist-lib.
//
// `base: "./"` keeps asset paths relative, so the same build works at the dev
// root (`/`) and deployed under a subpath (Caddy's `/multi`).
export default defineConfig({
  base: "./",
  server: { port: 5181 },
  optimizeDeps: { exclude: ["@forevka/wordcanvas"] },
  build: {
    target: "es2022",
    rollupOptions: { external: ["@forevka/wordcanvas"] },
  },
});
