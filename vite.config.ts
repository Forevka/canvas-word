import { defineConfig } from "vite";

export default defineConfig({
  // main.ts uses top-level await (font readiness gate, ?docx= import), which the
  // default "widely available" target rejects. Every browser that can run this
  // editor (canvas, Intl.Segmenter, module workers) supports ES2022 anyway.
  build: { target: "es2022" },
});
