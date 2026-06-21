import { defineConfig } from "vite";

// Tauri desktop frontend. Mirrors examples/embed-live: @forevka/wordcanvas is a
// PRE-BUILT, code-split ES bundle that loads its own editor app + import/export
// workers via `new URL(..., import.meta.url)`, so a downstream bundler must NOT
// re-process it. Keep it EXTERNAL:
//   - dev   → `optimizeDeps.exclude` serves the bundle as-is from node_modules.
//   - build → `rollupOptions.external` leaves the bare import in the output; the
//             index.html import map resolves it to the copied dist-lib at runtime
//             (see scripts/copy-lib.mjs).
//
// `base: "./"` keeps asset paths relative so the bundle resolves under Tauri's
// custom protocol origin (the editor's own workers/fonts are already relative —
// see frontend/vite.config.ts — this keeps OUR shell relative too).
const host = process.env.TAURI_DEV_HOST;

export default defineConfig({
  base: "./",
  // Tauri expects a fixed dev server it can point its webview at.
  clearScreen: false,
  server: {
    port: 1420,
    strictPort: true,
    // For mobile/LAN HMR set TAURI_DEV_HOST; otherwise Vite's defaults apply.
    ...(host ? { host, hmr: { protocol: "ws", host, port: 1421 } } : {}),
    // src-tauri changes are watched by the Rust side, not Vite.
    watch: { ignored: ["**/src-tauri/**"] },
  },
  optimizeDeps: { exclude: ["@forevka/wordcanvas"] },
  build: {
    // es2022 for the top-level await in the editor bundle.
    target: "es2022",
    rollupOptions: { external: ["@forevka/wordcanvas"] },
  },
  // Surface Tauri's build-time env vars to the frontend if ever needed.
  envPrefix: ["VITE_", "TAURI_ENV_"],
});
