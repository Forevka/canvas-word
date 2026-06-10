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

In a real app installed from npm, drop the import map and let your bundler resolve
the bare specifier:

```js
import { WordCanvas } from "@forevka/wordcanvas";
new WordCanvas({ container: document.getElementById("editor") });
```
