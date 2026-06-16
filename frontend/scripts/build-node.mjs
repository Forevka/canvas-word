// Node/server distribution build for @forevka/wordcanvas.
//
// The editor (`.` / `./builder`) is a browser bundle built by Vite. The headless
// pipelines, by contrast, are isomorphic and meant to run on a Node server for
// document import/export/TOC work. We bundle them with esbuild (not Vite) on
// purpose: esbuild leaves `new URL(`./fonts/${file}`, import.meta.url)` untouched,
// so measureHost's `node:fs` read resolves against the real ./fonts dir we copy
// next to the output — whereas Vite rewrites that expression into an inlined asset
// map (great for the browser, wrong for a Node fs read).
//
// Output: dist-node/{import,export,measure,recalc-docx,generate-toc}.js (+ shared
// chunks) and dist-node/fonts/*.ttf. Heavy deps (pdfkit, fontkit, fflate) are
// bundled so the package stays self-contained / zero runtime deps; Node builtins
// stay external (platform: "node").

import { build, context } from "esbuild";
import { cp, mkdir, rm } from "node:fs/promises";
import { fileURLToPath } from "node:url";

const here = (p) => fileURLToPath(new URL(p, import.meta.url));
const WATCH = process.argv.includes("--watch");

const OUT = here("../dist-node");
const FONTS_SRC = here("../src/export/shared/fonts");
const SHARED = here("../../shared/src/index.ts");

// Public Node entry points → source modules (mirrors package.json `exports`).
const ENTRIES = {
  import: here("../src/import/docx/pipeline.ts"),
  export: here("../src/export/pipeline.ts"),
  measure: here("../src/export/shared/measureHost.ts"),
  "recalc-docx": here("../src/recalc/patchTocDocx.ts"),
  "generate-toc": here("../src/recalc/generateTocDocx.ts"),
};

await rm(OUT, { recursive: true, force: true });

// Bundled metric-clone fonts: measureHost reads ./fonts/<file> relative to the
// emitted module via node:fs, so they must sit next to the output.
async function copyFonts() {
  await mkdir(OUT, { recursive: true });
  await cp(FONTS_SRC, new URL("../dist-node/fonts/", import.meta.url), { recursive: true });
}

const options = {
  entryPoints: ENTRIES,
  outdir: OUT,
  platform: "node",
  format: "esm",
  target: "node20",
  bundle: true,
  // ESM code-splitting so the shared layout engine / model / measure code is one
  // chunk instead of duplicated across the five entries.
  splitting: true,
  sourcemap: true,
  // @cw/shared is a private workspace package (raw TS) — resolve it to source so
  // its model/ops code is inlined into the bundle (it is never published).
  alias: { "@cw/shared": SHARED },
  // pdfkit/fontkit/fflate are bundled (zero runtime deps). Node builtins are
  // external automatically under platform "node".
  //
  // pdfkit is CJS and does `require("stream"|"zlib"|...)`. In ESM output esbuild's
  // require shim throws unless a real `require` is in scope, so synthesize one from
  // import.meta.url (also exposes __dirname for any CJS dep that reads it).
  banner: {
    js: [
      'import { createRequire as __createRequire } from "node:module";',
      'import { fileURLToPath as __fileURLToPath } from "node:url";',
      'import { dirname as __pathDirname } from "node:path";',
      "const require = __createRequire(import.meta.url);",
      "const __filename = __fileURLToPath(import.meta.url);",
      "const __dirname = __pathDirname(__filename);",
    ].join("\n"),
  },
  logLevel: "info",
};

if (WATCH) {
  const ctx = await context(options);
  await copyFonts();
  await ctx.watch();
  console.log("[build-node] watching dist-node/ (fonts copied once)");
} else {
  await build(options);
  await copyFonts();
  console.log("[build-node] wrote dist-node/ + fonts");
}
