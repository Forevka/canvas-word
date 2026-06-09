import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";
import { nodePolyfills } from "vite-plugin-node-polyfills";

// The export worker bundles pdfkit, which expects Node globals/builtins
// (process, buffer, stream, zlib, util, assert). Polyfill them for the BROWSER
// — both dev (`serve`) and prod (`build`). They must NOT be active under vitest,
// which runs on real Node where a faked global Buffer breaks pdfkit's
// `Buffer.isBuffer` checks. vitest sets process.env.VITEST.
const underVitest = process.env.VITEST !== undefined;

// `vite build --mode lib` produces the embeddable @cw/frontend package (entry
// src/wordcanvas.ts) into dist-lib; the default build produces the dev app
// (index.html) into dist.
export default defineConfig(({ mode }) => {
  const lib = mode === "lib";
  return {
    // Resolve the shared document core (@cw/shared) straight to its TS source so
    // dev/build/vitest all work without relying on a workspace symlink in
    // node_modules. Mirrors the `paths` entry in tsconfig.json.
    resolve: {
      alias: {
        "@cw/shared": fileURLToPath(new URL("../shared/src/index.ts", import.meta.url)),
      },
    },
    // ES2022: top-level await (font readiness gate) + every browser that runs the
    // editor (canvas, Intl.Segmenter, module workers) supports it.
    build: lib
      ? {
          target: "es2022",
          outDir: "dist-lib",
          emptyOutDir: true,
          lib: {
            entry: fileURLToPath(new URL("src/wordcanvas.ts", import.meta.url)),
            formats: ["es"],
            fileName: "wordcanvas",
          },
        }
      : {
          target: "es2022",
          // Multi-page app build: the default editor page + the /live online demo.
          rollupOptions: {
            input: {
              main: fileURLToPath(new URL("index.html", import.meta.url)),
              live: fileURLToPath(new URL("live/index.html", import.meta.url)),
            },
          },
        },
    plugins: underVitest
      ? []
      : [
          nodePolyfills({ include: ["buffer", "stream", "zlib", "util", "assert", "events", "process"] }),
          // Serve the /live page at the clean "/live" path (dev), not just "/live/".
          {
            name: "live-route",
            configureServer(server) {
              server.middlewares.use((req, _res, next) => {
                const url = req.url ?? "";
                const q = url.indexOf("?");
                const path = q === -1 ? url : url.slice(0, q);
                // Match the PATH only — a share link is /live?collab=… ; keep the
                // query so live.ts can read ?collab.
                if (path === "/live" || path === "/live/") {
                  req.url = "/live/index.html" + (q === -1 ? "" : url.slice(q));
                }
                next();
              });
            },
          },
        ],
    // ES-format workers so the export worker can code-split shared chunks (the
    // model/layout engine). The default IIFE format forbids code-splitting.
    worker: { format: "es" },
  };
});
