// Builds a representative document with the JS DocumentBuilder (directly, in Node)
// and exports it to a deterministic PDF — the reference for the C# `builddump` mode,
// which must build the IDENTICAL document via the typed wrapper. A byte-equal result
// proves the C# builder marshals to the same model the JS builder produces.
// Usage: node build-doc.mjs <outPdf>
import { readFileSync, readdirSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

const here = (p) => fileURLToPath(new URL(p, import.meta.url));
const code = readFileSync(here("../../dotnet/src/WordCanvas.ClearScript/assets/wordcanvas.clearscript.js"), "utf8");
(0, eval)(code);
const WC = globalThis.WordCanvas;

const fontsDir = here("../src/export/shared/fonts");
for (const f of readdirSync(fontsDir)) if (f.endsWith(".ttf")) WC.installFont(f, new Uint8Array(readFileSync(`${fontsDir}/${f}`)));
globalThis.__wc_fixedDate = 1_700_000_000_000; // deterministic timestamps

const b = WC.DocumentBuilder.create({ idSeed: "verify", pageSize: "Letter" });
b.style({ id: "Heading1", name: "Heading 1", type: "paragraph", char: { bold: true, fontSizePx: 24, color: "#1a1a2e" }, para: { spaceBeforePx: 18, spaceAfterPx: 8 } });
b.paragraph("canvas-word").align("center").font("Arial, sans-serif").fontSize(32).bold().color("#1a1a2e");
b.paragraph("a typed-builder parity check").align("center").italic().color("#5f6368");
b.paragraph("Section One").withStyle("Heading1");
b.paragraph().text("Plain, ").text("bold", { bold: true }).text(", a ").text("hyperlink", { link: "https://forevka.dev", color: "#0b57d0", underline: true }).text(", and page ").pageField().text(".");
b.table([["Feature", "Status"], ["Tables", "ok"], ["Lists", "ok"]], { headerRow: true });
b.bulletList(["bullet one", "bullet two"]);
b.numberedList(["number one", "number two"]);
b.paragraph("Body paragraph. " + "lorem ".repeat(20));
b.header((h) => { h.paragraph().text("canvas-word — parity"); });
b.footer((f) => { f.paragraph().align("center").text("Page ").pageField().text(" of ").numPagesField(); });

const doc = b.build();
const outPdf = process.argv[2] ?? here("../../out-js.pdf");
const pdf = await WC.exportPdf(doc);
writeFileSync(outPdf, Buffer.from(pdf));
const docx = await WC.exportDocx(doc);
// Derive a DISTINCT .docx path so a non-".pdf" outPdf can't clobber the PDF.
const outDocx = /\.pdf$/i.test(outPdf) ? outPdf.replace(/\.pdf$/i, ".docx") : `${outPdf}.docx`;
writeFileSync(outDocx, Buffer.from(docx));
console.log(`js build → pdf ${pdf.length} bytes, docx ${docx.length} bytes`);
