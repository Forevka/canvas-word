// Quick sanity for the ClearScript IIFE bundle, run under Node (has Buffer/timers
// natively, so this validates the pipeline + bundling, NOT the V8 polyfills — the
// .NET smoke test does that). Loads the bundle, injects fonts, round-trips a docx.
import { readFileSync, readdirSync } from "node:fs";
import { fileURLToPath } from "node:url";

const here = (p) => fileURLToPath(new URL(p, import.meta.url));
const code = readFileSync(here("../../dotnet/src/WordCanvas.ClearScript/assets/wordcanvas.clearscript.js"), "utf8");
// eslint-disable-next-line no-eval
(0, eval)(code);
const WC = globalThis.WordCanvas;
if (!WC) throw new Error("globalThis.WordCanvas not set by bundle");

const fontsDir = here("../src/export/shared/fonts");
let n = 0;
for (const f of readdirSync(fontsDir)) {
  if (f.endsWith(".ttf")) {
    WC.installFont(f, new Uint8Array(readFileSync(`${fontsDir}/${f}`)));
    n++;
  }
}
console.log(`fonts installed: ${n}`);

const docx = new Uint8Array(readFileSync(here("../src/import/docx/__fixtures__/real-word-basic.docx")));
const h = WC.importDocx(docx);
console.log(`import: blocks=${h.blockCount} media=${h.mediaCount} warnings=${h.warnings.length}`);

const pdf = await WC.exportPdf(h.doc, h.images);
const pdfHdr = String.fromCharCode(...pdf.slice(0, 5));
console.log(`exportPdf: bytes=${pdf.length} header=${JSON.stringify(pdfHdr)}`);
if (pdfHdr !== "%PDF-") throw new Error("PDF header wrong");

const out = await WC.exportDocx(h.doc, h.images);
console.log(`exportDocx: bytes=${out.length} pk=${out[0] === 0x50 && out[1] === 0x4b}`);
const re = WC.importDocx(out);
console.log(`re-import: blocks=${re.blockCount} (orig ${h.blockCount})`);

// Builder smoke
const b = WC.DocumentBuilder.create();
b.paragraph("Hello from the builder").bold(true).end().paragraph("Second line");
const doc = b.build();
const bpdf = await WC.exportPdf(doc);
console.log(`builder->pdf: bytes=${bpdf.length}`);
console.log("SMOKE OK");
