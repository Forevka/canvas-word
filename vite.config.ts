import { defineConfig } from "vite";
import { nodePolyfills } from "vite-plugin-node-polyfills";

// The export worker bundles pdfkit, which expects Node globals/builtins
// (process, buffer, stream, zlib, util, assert). Polyfill them for the BROWSER
// — both dev (`serve`) and prod (`build`). They must NOT be active under vitest,
// which runs on real Node where a faked global Buffer breaks pdfkit's
// `Buffer.isBuffer` checks. vitest sets process.env.VITEST.
const underVitest = process.env.VITEST !== undefined;

export default defineConfig({
  // main.ts uses top-level await (font readiness gate, ?docx= import), which the
  // default "widely available" target rejects. Every browser that can run this
  // editor (canvas, Intl.Segmenter, module workers) supports ES2022 anyway.
  build: { target: "es2022" },
  plugins: underVitest
    ? []
    : [nodePolyfills({ include: ["buffer", "stream", "zlib", "util", "assert", "events", "process"] })],
  // ES-format workers so the export worker can code-split shared chunks (the
  // model/layout engine). The default IIFE format forbids code-splitting.
  worker: { format: "es" },
});
