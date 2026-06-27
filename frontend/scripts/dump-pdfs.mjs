// Dump a PDF per real .docx using the SAME ClearScript bundle, but run under Node.
// Comparing these against the C# `pdfdump` output isolates host determinism: the
// bundle, pdfkit, pako (bundled zlib), and fontkit are identical in both — only the
// JS host differs. Usage: node dump-pdfs.mjs <outDir>
import { readFileSync, readdirSync, writeFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";

const here = (p) => fileURLToPath(new URL(p, import.meta.url));
const outDir = process.argv[2] ?? here("../../out-node");
mkdirSync(outDir, { recursive: true });

const code = readFileSync(here("../../dotnet/src/WordCanvas.ClearScript/assets/wordcanvas.clearscript.js"), "utf8");
(0, eval)(code);
const WC = globalThis.WordCanvas;
globalThis.__wc_fixedDate = 1_700_000_000_000; // deterministic timestamps for byte-compare

const fontsDir = here("../src/export/shared/fonts");
for (const f of readdirSync(fontsDir)) if (f.endsWith(".ttf")) WC.installFont(f, new Uint8Array(readFileSync(`${fontsDir}/${f}`)));

const frontend = process.argv[3] ?? here("..");
for (const f of readdirSync(frontend)) {
  if (!f.endsWith(".docx")) continue;
  const h = WC.importDocx(new Uint8Array(readFileSync(`${frontend}/${f}`)));
  const pdf = await WC.exportPdf(h.doc, h.images);
  const name = f.replace(/\.docx$/, "");
  writeFileSync(`${outDir}/${name}.pdf`, Buffer.from(pdf));
  process.stdout.write(`wrote ${name}.pdf (${pdf.length} bytes)\n`);
}
