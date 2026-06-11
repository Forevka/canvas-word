import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";

// Like embed-live, the pre-built @forevka/wordcanvas editor bundle stays
// EXTERNAL (dev: optimizeDeps.exclude serves it as-is; build: rollup external +
// the import map in index.html). The /builder subpath is different: in the
// workspace it resolves to raw TS source, so dev transforms it like app code —
// which needs the @cw/shared alias the library's own config uses. The
// production build externalizes it too and the import map points at the copied
// dist-lib/builder.js.
export default defineConfig({
  base: "./",
  server: { port: 5181 },
  optimizeDeps: { exclude: ["@forevka/wordcanvas"] },
  resolve: {
    alias: {
      "@cw/shared": fileURLToPath(new URL("../../shared/src/index.ts", import.meta.url)),
    },
  },
  build: {
    target: "es2022",
    rollupOptions: { external: ["@forevka/wordcanvas", "@forevka/wordcanvas/builder"] },
  },
});
