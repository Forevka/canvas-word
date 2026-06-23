# Fonts: why we render & embed metric clones (not system fonts)

**If you touch anything font-related in this project, read this first.**

## TL;DR

The editor and both exporters (PDF, DOCX) **render, measure, and embed a fixed set
of bundled, freely-redistributable metric-clone fonts** instead of the proprietary
Windows/Office originals:

| Original (model keeps this name) | Clone we actually use | License |
|---|---|---|
| Calibri | Carlito | OFL |
| Cambria | Caladea | OFL |
| Georgia | Gelasio | OFL |
| Arial / Helvetica / Verdana / Tahoma / Segoe UI / Trebuchet MS | Arimo | Apache-2.0 |
| Times New Roman / Times / Garamond | **Times New Roman (genuine)** | ⚠ proprietary (Microsoft) |
| Courier New / Consolas / Monaco | Cousine | Apache-2.0 |
| (anything unmapped) | Arimo (fallback) | — |

> **Times New Roman is the real Microsoft font**, bundled at the project owner's
> request in place of the OFL "Tinos" clone (`frontend/src/export/shared/fonts/TimesNewRoman-*.ttf`,
> from <https://github.com/misuchiru03/font-times-new-roman>). Unlike the other
> faces it is **not** freely redistributable — see `fonts/LICENSES.md`. To restore
> a fully-redistributable build, swap it back for Tinos (Apache-2.0).

Single source of truth: **`frontend/src/fonts/clones.ts`** (`CLONE_OF`, `FONT_FILES`,
`TOOLBAR_FONTS`). Font files + their licenses: `frontend/src/export/shared/fonts/`.

## Why this decision was made

1. **Deterministic, cross-platform layout.** This editor is *page-accurate*: the
   on-screen pagination must be reproducible in the browser **and** on a Node
   backend (server-side export). System fonts like Calibri exist only on
   Windows/Office machines — they can't be used server-side at all, and even
   across browsers/OSes their metrics vary. Pinning to bundled fonts makes layout
   reproducible everywhere instead of depending on whatever the host has installed.

2. **Legal redistribution.** We cannot ship Microsoft's fonts. Carlito/Caladea/
   Gelasio (OFL) and Arimo/Tinos/Cousine (Apache-2.0) are freely redistributable
   and were specifically designed to be **metric-compatible** (same advance widths
   and line metrics) with Calibri/Cambria/Georgia/Arial/Times New Roman/Courier
   New — so documents lay out the same and look near-identical.

3. **Measurement must equal embedding.** The PDF exporter measures with the very
   same font it embeds (both via fontkit), so painted glyphs land exactly where
   the layout engine broke the lines. Measuring with a system font but embedding a
   clone (or vice-versa) would desync text positions.

## How it's wired (the model never changes — only rendering does)

The document model keeps the **original** family name (`"Calibri, serif"`). The
substitution happens only at render/measure/embed time:

- **`frontend/src/layout/metrics.ts` → `charStyleToFont`** maps the family to its clone, so
  every canvas font string the engine builds uses the clone. This is the one
  place the editor's measurement & painting pick up the clone.
- **`frontend/src/layout/metrics.ts` → `fontMetrics`** returns line-height ascent/descent
  from baked per-clone ratios (`CLONE_METRICS` in `frontend/src/fonts/clones.ts`), **not**
  from the canvas/fontkit context. This is what makes pagination identical across
  the editor and exporters (see "Pagination parity" below).
- **`frontend/src/export/shared/editorFonts.ts` → `loadEditorFonts`** loads the clone TTFs
  as `FontFace`s (under their clone names) on the main thread **before the first
  layout** (`main.ts`). Without this the editor canvas couldn't draw the clones.
- **`frontend/src/export/shared/measureHost.ts` + `fontkitContext.ts`** measure with
  fontkit over the same clones — in the export worker and on Node.
- **`frontend/src/export/shared/fontRegistry.ts` + `pdf/renderPdf.ts`** embed the clone
  faces into the PDF (pdfkit, subset-embedded).
- **DOCX (`frontend/src/export/docx/…`)** writes the **original** family name
  (`w:rFonts w:ascii`), so a `.docx` reopened in Word uses the real Calibri. The
  clone is a *rendering/measuring* concern, never stored in the document.

## Toolbar labelling

The font dropdown shows **"Original (Clone)"** — e.g. `Calibri (Carlito)` — so a
user can see the on-screen text is a metric-compatible substitute, not the literal
system font. Defined in `TOOLBAR_FONTS` (`frontend/src/fonts/clones.ts`). The selected
*value* is still the original family (what gets stored in the model).

## Pagination parity: editor === browser export === Node export

The editor and the exporters get text geometry from two different engines —
the editor measures through the browser's **2D canvas**, the exporters through
**fontkit** (pure JS, so it runs on Node with no DOM). The two halves are kept in
lock-step so a document paginates identically everywhere:

- **Widths / line breaking** — canvas and fontkit, over the *same* clone font
  files, already agree to sub-pixel (e.g. `"The quick brown fox"` measures 134.63
  in both), so line breaks are identical with no extra work.
- **Line heights** — these used to diverge: the browser's canvas `fontBoundingBox`
  and fontkit's font tables report slightly different ascent/descent for the same
  font, which over a long document tipped pagination by a page (a trailing empty
  paragraph spilling onto a new page). Fixed by computing `fontMetrics` from
  **baked per-clone ratios** (`CLONE_METRICS`) instead of either context's native
  values — so the editor and both exporters use the *same* line-height function on
  every platform. The ratios were measured from the browser's own canvas values at
  the 16px body size (an exact multiple), so the editor's appearance is unchanged.

Result (verified on the 8.9 MB sample report): **editor = browser export = Node
export = 80 pages.** If you change `CLONE_METRICS` or `charStyleToFont`, re-verify
this parity (lay the same doc out in the editor and via `runExport` and compare
page counts) — they must stay equal. See `EXPORT.md` for the export architecture.

## Custom fonts (embedder-supplied, at runtime)

Embedders can add their own fonts via the `fonts` constructor option without
touching the bundle:

```ts
new WordCanvas({
  fonts: {
    disableBuiltin: ["Calibri"],          // hide built-ins from the TOOLBAR only
    fonts: [{
      family: "Inter",                     // model + toolbar + render name
      faces: { regular: "https://…/Inter-Regular.ttf", bold: "https://…/Inter-Bold.ttf" },
      sizing: { ascent: 0.95, descent: 0.24 },   // REQUIRED — same role as CLONE_METRICS
    }],
  },
})
```

How it threads through (mirrors the clone path, so parity holds):

- **`frontend/src/fonts/customRegistry.ts`** is a global, additive overlay. The
  `clones.ts` resolvers consult it: `cloneFamilyFor` returns a custom family **as
  itself** (`substituted: false`, no Arimo fallback — it shadows a built-in of the
  same name), and `metricsFor` returns the caller's `sizing` instead of
  `CLONE_METRICS`. So a custom family renders/measures/embeds as itself with its
  supplied vertical metrics — the one knob that keeps the editor and both exporters
  paginating identically. **`sizing` is required for exactly this reason.**
- **Editor:** `editorFonts.ts` loads each face as a `FontFace` under the custom
  family name (missing bold/italic reuse the regular bytes — matching export).
- **Export:** the main thread fetches face bytes and passes them in the worker
  message (`ToExportWorker.fonts`); `runExport` calls `registerCustomFonts` +
  `registerCustomFontBytes` before layout; `resolveFont` keys custom faces by a
  synthetic `__custom__<family>-<Style>.ttf` name (both the fontkit-registry key and
  the pdfkit registration key). DOCX writes the custom family name verbatim
  (`w:rFonts w:ascii`).
- **Backend:** `backend/src/export/fontCache.ts` fetches + disk-caches faces by URL
  hash; `serverExport` reads the doc's saved config (`ChangeStore.getFontsConfig`,
  set at `POST /docs` time) or, for `/render.pdf`, a `fonts` multipart part.

Constraints (documented behavior, not bugs):

- **TTF/OTF only** — fontkit can't parse WOFF2 and the exporter must parse the exact
  bytes it embeds (WOFF2 is rejected at registration with a warning).
- **Missing bold/italic/boldItalic → regular** in BOTH editor and export
  (deterministic parity over faux-synthesis).
- The custom-font registry is **process-global** (like the rest of the font
  subsystem); the toolbar is per-instance. Two instances cannot define the **same
  family name** with **different** faces/sizing — latest registration wins, with a
  `console.warn`.
- Font hosts must allow **CORS** (editor `FontFace.load` + main-thread fetch). A URL
  that fails to load warns and falls back to a clone rather than hard-failing.

## Adding or changing a *bundled* font

1. Drop the four faces (`Family-Regular/Bold/Italic/BoldItalic.ttf`) into
   `frontend/src/export/shared/fonts/`. **Only ship OFL/Apache (or otherwise
   redistributable) fonts** and add them to `fonts/LICENSES.md`.
2. Update `CLONE_FAMILIES`, `CLONE_OF`, and `TOOLBAR_FONTS` in `frontend/src/fonts/clones.ts`.
   `FONT_FILES` derives automatically; the editor loader, fontkit measure host, and
   pdfkit embedder all read from there, so nothing else needs editing.
3. Run the export tests (`npx vitest run src/export/`, from `frontend/`) — the measure-host test
   asserts fontkit widths match pdfkit's embedded-font widths.
