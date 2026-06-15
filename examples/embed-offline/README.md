# Offline embed (no build step)

The smallest possible integration of `@forevka/wordcanvas`: a static HTML page
that mounts the editor in fully offline mode (no backend, no sync, no share),
using a browser **import map** to resolve the package — no bundler, no tooling.

## Run it

From the **repo root** (so `/node_modules/...` in the import map resolves):

```sh
npm install            # once, links the workspace packages
npm run build:lib      # produces frontend/dist-lib that the import map points at
npx serve .            # or: python -m http.server
```

Then open `http://localhost:3000/examples/embed-offline/` (port depends on your
static server).

## What to look at

- `index.html` — the import map mapping `@forevka/wordcanvas` to the built bundle.
- `app.js` — `import { WordCanvas } from "@forevka/wordcanvas"; new WordCanvas({ container })`.
- `app.js` also wires a **`resolveField`** callback. Open a `.docx` containing a
  developer-defined field — a paragraph whose `w:instrText` is something like
  ` MYCHART "sales-2026" ` — then right-click it and choose **Update Field
  (MYCHART)**. The engine calls `resolveField({ name, instruction })`, parses the
  OOXML it returns, and splices it in as the field's result. In production you'd
  forward that request to your backend (which renders the OOXML) instead of
  synthesizing it locally. Built-in `TOC` fields show **Update Field (TOC)** and
  regenerate locally (no backend).

In a real app installed from npm, drop the import map and let your bundler resolve
the bare specifier:

```js
import { WordCanvas } from "@forevka/wordcanvas";
new WordCanvas({ container: document.getElementById("editor") });
```
