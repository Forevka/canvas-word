import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";

// Like embed-live, the pre-built @forevka/wordcanvas editor bundle stays
// EXTERNAL (dev: optimizeDeps.exclude serves it as-is; build: rollup external +
// the import map in index.html). The /builder subpath is different: the package
// now publishes ./builder -> dist-lib/builder.js, but in DEV we want the raw TS
// source so changes show up without a lib rebuild — so we alias it to the
// frontend source (which needs the @cw/shared alias the library's own config
// uses). The alias is dev-only: the production build externalizes /builder and
// the import map points at the copied dist-lib/builder.js, so it must stay a
// bare specifier for rollup's `external` to match it.
export default defineConfig(({ command }) => ({
  base: "./",
  server: { port: 5181 },
  optimizeDeps: { exclude: ["@forevka/wordcanvas"] },
  resolve: {
    alias: {
      "@cw/shared": fileURLToPath(new URL("../../shared/src/index.ts", import.meta.url)),
      ...(command === "serve"
        ? {
            "@forevka/wordcanvas/builder": fileURLToPath(
              new URL("../../frontend/src/builder/index.ts", import.meta.url),
            ),
          }
        : {}),
    },
  },
  build: {
    target: "es2022",
    rollupOptions: { external: ["@forevka/wordcanvas", "@forevka/wordcanvas/builder"] },
  },
}));
