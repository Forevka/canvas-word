import { defineConfig } from "vite";

// Like embed-live, this only imports the @forevka/wordcanvas MAIN entry, so the
// config stays minimal — no @cw/shared alias, no /builder plumbing.
//
// The library is a pre-built, code-split ES bundle that loads its editor app +
// export/import workers via `new URL(..., import.meta.url)`, so a downstream
// bundler must NOT re-process it. Keep it EXTERNAL:
//   - dev   → `optimizeDeps.exclude` serves the bundle as-is from node_modules.
//   - build → `rollupOptions.external` leaves the bare import in the output; the
//             import map in index.html resolves it to the copied dist-lib at
//             runtime (see the `build` script's copy step + README).
//
// `base: "./"` keeps built asset paths relative, so the same build works at the
// dev root (`/`) and deployed under a subpath (Caddy's `/examples/constructor`).
export default defineConfig({
  base: "./",
  server: { port: 5183 },
  optimizeDeps: { exclude: ["@forevka/wordcanvas"] },
  build: {
    target: "es2022",
    rollupOptions: { external: ["@forevka/wordcanvas"] },
  },
});
