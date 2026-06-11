import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";
import { nodePolyfills } from "vite-plugin-node-polyfills";

// The export worker bundles pdfkit, which expects Node globals/builtins
// (process, buffer, stream, zlib, util, assert). Polyfill them for the BROWSER
// — both dev (`serve`) and prod (`build`). They must NOT be active under vitest,
// which runs on real Node where a faked global Buffer breaks pdfkit's
// `Buffer.isBuffer` checks. vitest sets process.env.VITEST.
const underVitest = process.env.VITEST !== undefined;

// `vite build --mode lib` produces the embeddable @forevka/wordcanvas package
// (entry src/wordcanvas.ts) into dist-lib; the default build produces the dev app
// (index.html) into dist.
export default defineConfig(({ mode }) => {
  const lib = mode === "lib";
  return {
    // Relative asset/worker URLs (`new URL("./assets/…", import.meta.url)`) so the
    // bundle is relocatable: embedders mount dist-lib under an arbitrary subpath
    // (e.g. /live/wordcanvas/) and the editor's import/export workers still
    // resolve against their own location. With the default base "/", worker URLs
    // become origin-absolute (/assets/…) and 404 under any subpath — the worker
    // then fails to load with an opaque error. The dev app keeps the default base.
    base: lib ? "./" : "/",
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
            // Multi-entry (ES-only, which we already are): the editor package
            // plus the standalone builder (@forevka/wordcanvas/builder), which
            // pulls in the import pipeline but none of the editor UI.
            entry: {
              wordcanvas: fileURLToPath(new URL("src/wordcanvas.ts", import.meta.url)),
              builder: fileURLToPath(new URL("src/builder/index.ts", import.meta.url)),
            },
            formats: ["es"],
          },
        }
      : {
          target: "es2022",
          // Single-page dev app: the editor harness (index.html). The online +
          // collaboration demo lives in examples/embed-live, which consumes the
          // built @forevka/wordcanvas package the way a real embedder would.
        },
    plugins: underVitest
      ? []
      : [nodePolyfills({ include: ["buffer", "stream", "zlib", "util", "assert", "events", "process"] })],
    // ES-format workers so the export worker can code-split shared chunks (the
    // model/layout engine). The default IIFE format forbids code-splitting.
    worker: { format: "es" },
  };
});
