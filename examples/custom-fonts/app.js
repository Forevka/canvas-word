// Custom fonts demo (OFFLINE, no build step).
//
// The `fonts` constructor option registers an embedder-supplied font, loaded from
// URLs at runtime, and can hide built-ins from the toolbar. Here we self-host PT
// Serif (OFL, in ./fonts) so the demo works fully offline — but the URLs could be
// any CORS-enabled host (a CDN, your asset server). The REQUIRED `sizing`
// (ascent/descent as fractions of em) is what keeps the editor and the PDF/DOCX
// exporters paginating identically — same role as the built-in clones' baked
// metrics. PT Serif's values come from its font metrics (1039/-286 per 1000 em).
import { WordCanvas } from "@forevka/wordcanvas";
import { DocumentBuilder } from "@forevka/wordcanvas/builder";

const FONTS = {
  // Hide a built-in from the toolbar (it stays loaded + resolvable for any .docx
  // that references it — this only trims the font dropdown).
  disableBuiltin: ["Calibri"],
  fonts: [
    {
      family: "PT Serif", // stored in the model, shown in the toolbar, the render name
      faces: {
        regular: "./fonts/PTSerif-Regular.ttf",
        bold: "./fonts/PTSerif-Bold.ttf",
        italic: "./fonts/PTSerif-Italic.ttf",
        boldItalic: "./fonts/PTSerif-BoldItalic.ttf",
      },
      sizing: { ascent: 1.039, descent: 0.286 },
    },
  ],
};

// First-load loading bar (see #loader in index.html).
const loader = document.getElementById("loader");
const bar = loader.querySelector(".bar");
const label = loader.querySelector(".label");
const PHASE_LABEL = { bundle: "Loading editor…", fonts: "Loading fonts…", ready: "Ready" };

const panel = document.getElementById("panel");
const status = document.getElementById("status");

const editor = new WordCanvas({
  container: document.getElementById("editor"),
  fonts: FONTS,
  onLoadProgress: ({ phase, percent }) => {
    bar.style.width = `${Math.round(percent * 100)}%`;
    label.textContent = PHASE_LABEL[phase];
    if (phase === "ready") loader.classList.add("done");
  },
});

// Build a document that actually USES the custom font, so it's visible on open and
// proves the faces render. `.font("PT Serif")` sets the run family; bold()/italic()
// pick the matching face (a missing face would fall back to regular).
function demoDoc() {
  const b = DocumentBuilder.create();
  b.paragraph("Custom Fonts in WordCanvas").withStyle("Heading1").font("PT Serif");
  b.paragraph(
    "This whole document is set in PT Serif — a custom OFL font supplied through the " +
      "fonts config and loaded from ./fonts at runtime, not bundled with the editor.",
  ).font("PT Serif");
  b.paragraph("Regular, bold, and italic all use the real face files:").font("PT Serif");
  b.paragraph("PT Serif Bold — a real bold face, not faux-synthesized.").font("PT Serif").bold();
  b.paragraph("PT Serif Italic — a real italic face.").font("PT Serif").italic();
  b.paragraph("PT Serif Bold Italic — the fourth face.").font("PT Serif").bold().italic();
  b.paragraph(
    "For contrast, this line uses the built-in Arial (Arimo) clone. Open the Font " +
      'dropdown: "PT Serif" is listed and "Calibri" is hidden via disableBuiltin.',
  ).font("Arial");
  return b.build();
}

editor.on("ready", () => {
  panel.hidden = false;
  try {
    editor.setDocument(demoDoc());
  } catch (e) {
    console.error("failed to build demo document", e);
  }
  // Diagnostics for E2E / quick poking.
  window.__wc = editor;
  window.__fontReady = document.fonts.check("16px 'PT Serif'");
  window.__ready = true;
  console.log("WordCanvas ready — PT Serif loaded:", window.__fontReady);
});

// Export to PDF — proves the custom font embeds. Downloads the file and reports the
// byte size (the PDF subset-embeds the PT Serif faces actually used).
document.getElementById("export").addEventListener("click", async () => {
  status.textContent = "Exporting…";
  try {
    const blob = await editor.exportPdf();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "custom-fonts-demo.pdf";
    a.click();
    URL.revokeObjectURL(url);
    status.textContent = `Exported PDF (${Math.round(blob.size / 1024)} KB) with PT Serif embedded.`;
  } catch (e) {
    status.textContent = `Export failed: ${e?.message ?? e}`;
    console.error(e);
  }
});
