# Export (PDF + DOCX)

The editor exports the document model to **PDF** (page-accurate) and **DOCX**
(hand-rolled OOXML). Both run **isomorphically**: the same pure pipeline executes
in a browser Web Worker *and* on a Node backend (no DOM), tested under vitest.
The feature is additive under `frontend/src/export/`, plus one small seam in
`frontend/src/layout/metrics.ts`.

```
exportDocument(doc, "pdf"|"docx")   main thread: resolve image bytes, post to worker
  └─ worker.ts ─ runExport()        pure, DOM-free (also the Node entry)
       ├─ pdf/renderPdf.ts          layout engine → LayoutTree → pdfkit
       └─ docx/writeDocx.ts         model → OOXML parts → fflate zip
```

## PDF — page-accurate, by reusing the layout engine

PDF is **not** a reflow. `renderPdf` runs the editor's own
`createLayoutEngine().layout(doc)` to get the `LayoutTree` (absolute page
geometry), then `pdf/paintBlock.ts` draws each page with pdfkit — a
constant-for-constant inverse of `frontend/src/paint/renderer.ts` (same baseline
formula, sub/super shifts, underline/strike offsets, default grid color,
footnote-rule width, leader dashes). A PDF page matches the canvas
pixel-for-pixel (modulo metric-clone glyph shapes). The model is CSS px (96dpi);
each PDF page is sized in points and scaled by 72/96, so the painter keeps
drawing in document px.

### Running the layout engine without a DOM (the crux)

The engine measures text through a 2D canvas in two spots: `metrics.ts` and
`@chenglou/pretext` (`getMeasureContext`). `shared/measureHost.ts` makes both
DOM-free by routing them through a **fontkit-backed shim** (`shared/fontkitContext.ts`)
over the bundled fonts — the **same** path in the browser worker and on Node, so an
exported PDF is identical across environments (no system fonts anywhere). It also
loads the bundled faces for pdfkit to embed.

The editor renders the **same bundled fonts** (`charStyleToFont` maps families to
clones, `shared/editorFonts.ts` loads them as `FontFace`s), and **`fontMetrics`
computes line heights from baked per-font ratios** rather than the live
canvas/fontkit context, so the editor, the browser export, and the Node export all
paginate identically. See `FONTS.md` ("Pagination parity") for the why and how.

### pdfkit JPEG colorspace fix

pdfkit 0.19's JPEG parser reads the first component's *id* instead of the
component *count*, tagging every standard color JPEG `DeviceGray` (renders grayed
and garbled). `pdf/paintBlock.ts` parses the real SOF component count and
overrides `colorSpace` (via pdfkit's `openImage`) before placing the image.

### Fonts

`shared/fonts/` ships the bundled faces (each R/B/I/BI): Carlito↔Calibri,
Caladea↔Cambria, Gelasio↔Georgia, Arimo↔Arial, Cousine↔Courier New, plus the
**genuine Times New Roman** (bundled on request — proprietary; see `FONTS.md` and
`fonts/LICENSES.md`). The family→clone map and substitution live in
`frontend/src/fonts/clones.ts` (`cloneFamilyFor`), shared with the editor; `Arimo` is the
fallback (emits a `font-substituted` warning). `shared/fontRegistry.ts` loads/
resolves them via fontkit. Node reads the files with `fs`; the browser worker
`fetch`es them (Vite emits them as assets). **Full rationale: `FONTS.md`.**

## DOCX — hand-rolled OOXML, the inverse of the importer

`writeDocx` emits `document.xml`, `styles.xml`, `numbering.xml`, `footnotes.xml`,
`settings.xml`, header/footer parts, `[Content_Types].xml`, and `.rels`, then zips
with `fflate`. Unit inverses live in `frontend/src/export/units.ts` (px→twips ×15,
→half-pt ×1.5, →EMU ×9525, border →eighth-pt ×6). Runs carry full direct
formatting (every toggle explicit on/off) so a paragraph's `w:pStyle` can't leak
run props back through the cascade on re-import. Row spans re-synthesize the
`w:vMerge="continue"` cells the importer drops; list ids are remapped to integer
`numId`s shared between `numbering.xml` and `w:numPr`.

**Correctness oracle:** the round-trip `writeDocx → runImport → compare`. The real
report re-imports to the **exact same block count** (1031 → 1031).

## Performance (real 8.9MB, 81-page report, Node)

import ~220ms · PDF ~1.2s (9.5MB, 81pp) · DOCX ~0.4s (9.1MB) · docx re-import 1031→1031 blocks.

## Saving from an embedder (route exports to your own pipeline)

An embedder who wants a **Save** button that ships the file to their own backend
has two entry points on the `WordCanvas` package API. Both reuse the in-worker
pipeline above (no extra bundle, no `installMeasureHost()` dance), and both bake
the track-changes overlay to the original baseline like the toolbar's Export:

- **`exportDocx()` / `exportPdf()`** on the instance (and its `EditorHandle`) each
  resolve to a `Blob`. Wire your own button and `POST` the result anywhere. These
  work even when the ribbon is hidden (`view.toolbar: false` / `readonly`).

  ```ts
  const ed = new WordCanvas({ container });
  saveBtn.onclick = async () => {
    const blob = await ed.exportDocx();
    await fetch("/api/documents", { method: "POST", body: blob });
  };
  ```

- **`onSave`** constructor option. When set, the toolbar's **Export (PDF / DOCX)**
  buttons hand the produced file to your callback (`SaveEvent`: `blob`, raw
  `bytes`, `format`, export `warnings`) instead of triggering a browser download.
  Return a promise to keep the UI responsive while you upload.

  ```ts
  new WordCanvas({
    container,
    onSave: async ({ blob, format }) => {
      await fetch(`/api/documents?format=${format}`, { method: "POST", body: blob });
    },
  });
  ```

Both sit on top of `exportDocument(doc, format)`; `getDocument()` (the raw model
snapshot) and the headless `@forevka/wordcanvas/export` `runExport()` remain
available for fully custom flows.

To **hide the built-in buttons** entirely (an embedder driving export from its own
UI), set `view.exportPdf: false` and/or `view.exportDocx: false` — the Export group
disappears when both are off, while `exportDocx()` / `exportPdf()` keep working.

## Bundling

The export worker is an **ES-format** worker (`vite.config.ts` `worker.format`)
so it can code-split the shared engine. pdfkit wants Node builtins, polyfilled for
the **production browser build only** (`nodePolyfills`, `command === "build"`) —
never on Node (a faked global `Buffer` breaks pdfkit). `font: false` on the
`PDFDocument` skips pdfkit's eager `Helvetica.afm` `fs` read, the one remaining
Node-only dependency, so the browser worker never touches `fs`.

## Lossy / known limitations

- **CJK & complex scripts**: only Latin fonts are bundled, so CJK renders as tofu
  in the PDF (same east-asian-fonts gap as the importer; real reports are Latin).
- **Pagination parity**: editor, browser export, and Node export paginate
  identically — they share `charStyleToFont` + baked `fontMetrics` over the same
  bundled fonts (see `FONTS.md`). Not byte-identical: the browser and Node PDF
  *bytes* still differ slightly (pdfkit timestamp + zlib implementation), but the
  layout/page count is the same.
- **Symbol/decorative fonts** (e.g. Symbol, Wingdings) fall back → `font-substituted`.
- **Per-section bands on mid-document section breaks** aren't emitted (the
  document section's bands are used) — mirrors the importer's `section-bands-flattened`.
- **Column breaks** export best-effort for Word; not part of the round-trip oracle.
- **Browser PDF runtime** is verified end-to-end via Playwright (loads the editor,
  exports, checks page count & image colorspaces); the Node path is fully
  test-covered. DOCX is pure-JS and identical in both.
