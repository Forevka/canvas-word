# embed-multi

Mounts **several `WordCanvas` editors on a single page**, each fully independent
(own document, undo history, ribbon, ruler, and selection). Use the header buttons
to add or remove editors at runtime — removing one calls `destroy()`, which tears
down its chrome and its body-level floating panels (find bar, image toolbar, …)
without touching the others.

This works because the library's chrome is **class-scoped** under a per-instance
`.wordcanvas-root` (no global ids), and each instance owns its off-container
artifacts and global listeners (cleaned up on `destroy()`).

```ts
import { WordCanvas } from "@forevka/wordcanvas";

const a = new WordCanvas({ container: document.getElementById("a")! });
const b = new WordCanvas({ container: document.getElementById("b")! });
// …later
a.destroy(); // b keeps working
```

## Run

The example imports the **built** `dist-lib` bundle, so build the library first:

```bash
# from the repo root
npm run build:lib --workspace @forevka/wordcanvas
npm run dev --workspace @cw/example-embed-multi
```

`npm run build --workspace @cw/example-embed-multi` produces a self-contained
`dist/` (app + a copied `./wordcanvas` bundle). Deployed at `/multi`.

## Connect an AI agent (WebMCP)

Each editor sets `agentTools: { name: "editor<N>" }`, exposing it to AI agents over
[WebMCP](https://webmcp.dev) (`navigator.modelContext`). The **`name`** namespaces
each instance's tools (e.g. `editor0_get_document`, `editor1_replace_text`) so an
agent — connected via the WebMCP browser extension / Chrome DevTools MCP — can
target a specific editor among the several on the page. `destroy()` (Remove last)
unregisters that instance's tools.
