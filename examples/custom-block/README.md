# Custom block type

[![Open in StackBlitz](https://developer.stackblitz.com/img/open_in_stackblitz.svg)](https://stackblitz.com/github/Forevka/canvas-word/tree/main/examples/custom-block)

Add a **new document block type** — one that draws itself on the page canvas —
through the **`registerBlockType`** registry, instead of forking the layout
engine and paint code. The demo registers a `barchart` block and inserts it into
the document, where it paginates and renders like any built-in block. A static,
no-build page (browser import map), like [`embed-offline`](../embed-offline).

## Run it

From the **repo root** (so `/node_modules/...` in the import map resolves):

```sh
npm install
npm run build:lib
npx serve .            # or: python -m http.server
```

Then open `http://localhost:3000/examples/custom-block/`.

## What to look at — `app.js`

```js
import { registerBlockType } from "@forevka/wordcanvas";

registerBlockType({
  type: "barchart",
  // Atomic: return the block's height for the available content width.
  measure: (data, ctx) => ({ height: 168 }),
  // Draw on the page canvas — ctx is translated to the block's top-left, so you
  // draw in local [0, width] × [0, height] document px.
  paint: (ctx, box, data) => { /* draw bars from data.values */ },
  // Optional: toOOXML(data) => "<w:p>…</w:p>" for .docx export; omitted here.
});
```

A custom block in the model is just:

```js
{ kind: "custom", id, revision: 0, customType: "barchart", data: { title, values } }
```

Insert it wherever blocks go (the demo splices it into `doc.blocks` and calls
`handle.setDocument`).

## The block contract

- **Atomic** — measures to a single box and places whole, like an image or
  equation. It stays out of line-breaking and lays out anywhere blocks go (body,
  table cells, header/footer bands).
- **A first-class object** — **click to select** (a plain selection frame),
  **Delete** or **right-click ▸ Delete** to remove, with full **undo/redo**. Try
  it: click the chart, then press Delete, then Ctrl/Cmd+Z.
- **JSON-serializable `data`** — so snapshot serialize/paste is free (the data is
  deep-cloned on paste).
- **Canvas-drawn** — `paint` receives a `CanvasRenderingContext2D` already
  translated + clipped to the block's box; it draws only (never measures text).
- **PDF renders your block** — the same `paint` is replayed through a
  Canvas2D→pdfkit vector shim, so the block appears in the PDF (crisp vectors),
  headlessly. No extra wiring.
- **`.docx` is lossy unless you provide `toOOXML`** — a custom block has no native
  OOXML, so by default `.docx` export emits a placeholder paragraph and reports a
  `custom-block-dropped` warning. Supply `toOOXML(data)` to control the DOCX
  output. Try the **Export** button and watch the console.

If a block's `customType` isn't registered, the editor paints a visible dashed
placeholder (so a missing `registerBlockType` is obvious, not an invisible gap).

**By design**, a custom block's content is *drawn*, not text-edited — it has no
internal caret (that's what paragraphs and content controls are for), and its size
is owned by `measure()` rather than interactive resize.
