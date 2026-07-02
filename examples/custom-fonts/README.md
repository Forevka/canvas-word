# Custom fonts (no build step)

[![Open in StackBlitz](https://developer.stackblitz.com/img/open_in_stackblitz.svg)](https://stackblitz.com/github/Forevka/canvas-word/tree/main/examples/custom-fonts)

Embed `@forevka/wordcanvas` with the **`fonts`** option: register your own font,
loaded from URLs at runtime, with per-style faces and the required `sizing`
metrics — and optionally hide built-ins from the toolbar. This example is fully
offline (no backend) and self-hosts [PT Serif](https://fonts.google.com/specimen/PT+Serif)
(OFL — see `fonts/OFL.txt`) so it works with no network; the URLs could just as
well point at a CORS-enabled CDN.

```js
new WordCanvas({
  container,
  fonts: {
    disableBuiltin: ["Calibri"],            // hide a built-in from the toolbar only
    fonts: [{
      family: "PT Serif",                    // model name + toolbar + render name
      faces: {
        regular:    "./fonts/PTSerif-Regular.ttf",
        bold:       "./fonts/PTSerif-Bold.ttf",
        italic:     "./fonts/PTSerif-Italic.ttf",
        boldItalic: "./fonts/PTSerif-BoldItalic.ttf",
      },
      sizing: { ascent: 1.039, descent: 0.286 },   // REQUIRED — keeps editor == export pagination
    }],
  },
});
```

## Run it

From the **repo root** (so `/node_modules/...` in the import map resolves):

```sh
npm install            # once, links the workspace packages
npm run build:lib      # produces frontend/dist-lib that the import map points at
npx serve .            # or: python -m http.server
```

Then open `http://localhost:3000/examples/custom-fonts/` (port depends on your
static server).

## What to look at

- `app.js` — the `fonts` config, and a `DocumentBuilder` document that sets every
  paragraph in `PT Serif` (regular/bold/italic/bold-italic) so all four real faces
  render on open. The **Export PDF** button proves the custom font subset-embeds.
- Open the **Font dropdown** in the ribbon: `PT Serif` is listed (a custom family),
  and `Calibri` is gone (hidden via `disableBuiltin` — but still resolvable, so a
  loaded `.docx` that uses Calibri still renders).
- `fonts/` — the self-hosted TTF faces + the OFL license.

## Notes

- **TTF/OTF only.** WOFF2 is rejected (the exporter must parse the exact bytes it
  embeds, and fontkit can't read WOFF2). A missing bold/italic/bold-italic face
  falls back to `regular` in both the editor and the exporters.
- **`sizing` is required.** It's the ascent/descent (fractions of em) the layout
  engine uses for line height, on every platform — so the editor, the browser
  export, and a headless Node export paginate identically.
- Font hosts must allow **CORS** (the editor loads faces via `FontFace`, and export
  fetches the bytes on the main thread). Self-hosting same-origin sidesteps this.
- See [`FONTS.md`](../../FONTS.md) for the full design (global overlay, parity).

In a real app installed from npm, drop the import map and let your bundler resolve
the bare specifiers:

```js
import { WordCanvas } from "@forevka/wordcanvas";
```
