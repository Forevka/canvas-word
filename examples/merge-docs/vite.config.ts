import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";

// The pre-built @forevka/wordcanvas editor bundle (the `.` entry, used for the
// live preview) stays EXTERNAL — dev serves it as-is (optimizeDeps.exclude), the
// production build externalizes it and the index.html import map points at the
// copied dist-lib. The `/builder` + `/query` subpaths are aliased to the frontend
// SOURCE in DEV so this example runs against the current tree without a lib rebuild
// (and demonstrates the merge API even before it is published); the production
// build externalizes them too, resolved via the import map.
export default defineConfig(({ command }) => ({
  base: "./",
  server: { port: 5183 },
  optimizeDeps: { exclude: ["@forevka/wordcanvas"] },
  resolve: {
    alias: {
      "@cw/shared": fileURLToPath(new URL("../../shared/src/index.ts", import.meta.url)),
      ...(command === "serve"
        ? {
            "@forevka/wordcanvas/builder": fileURLToPath(new URL("../../frontend/src/builder/index.ts", import.meta.url)),
            "@forevka/wordcanvas/query": fileURLToPath(new URL("../../frontend/src/query.ts", import.meta.url)),
          }
        : {}),
    },
  },
  build: {
    target: "es2022",
    rollupOptions: { external: ["@forevka/wordcanvas", "@forevka/wordcanvas/builder", "@forevka/wordcanvas/query"] },
  },
}));
