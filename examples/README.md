# Examples

Standalone examples of embedding [`@forevka/wordcanvas`](../frontend) as a library,
each consuming the package by name the way a third-party integrator would.

| Example | What it shows | Tooling | Try |
| --- | --- | --- | --- |
| [`embed-offline`](./embed-offline) | Smallest possible integration — fully local editor (no backend), via a browser import map | None (static HTML) | [![Open in StackBlitz](https://developer.stackblitz.com/img/open_in_stackblitz.svg)](https://stackblitz.com/github/Forevka/canvas-word/tree/main/examples/embed-offline) |
| [`embed-live`](./embed-live) | Online editing + live collaboration + presence + share links (migrated from the old `/live` page) | Vite | [![Open in StackBlitz](https://developer.stackblitz.com/img/open_in_stackblitz.svg)](https://stackblitz.com/github/Forevka/canvas-word/tree/main/examples/embed-live) |
| [`playground`](./playground) | Programmatic `DocumentBuilder` playground — edit code, see the document rebuild live | Vite | [![Open in StackBlitz](https://developer.stackblitz.com/img/open_in_stackblitz.svg)](https://stackblitz.com/github/Forevka/canvas-word/tree/main/examples/playground) |
| [`query-edit`](./query-edit) | Headless query + edit API (`/query`) — find/sections/SDTs/pages + `DocumentEditor` (fill controls, find-replace, move, undo/redo) | Vite | [![Open in StackBlitz](https://developer.stackblitz.com/img/open_in_stackblitz.svg)](https://stackblitz.com/github/Forevka/canvas-word/tree/main/examples/query-edit) |
| [`merge-docs`](./merge-docs) | Merge/append documents (`/query`) — fold parts with `DocumentEditor.append` / `mergeAll`, then apply a per-section footer (table + logo) with `setSectionFooter` | Vite | [![Open in StackBlitz](https://developer.stackblitz.com/img/open_in_stackblitz.svg)](https://stackblitz.com/github/Forevka/canvas-word/tree/main/examples/merge-docs) |
| [`embed-multi`](./embed-multi) | Multiple independent editors on one page; add/remove instances at runtime | Vite | [![Open in StackBlitz](https://developer.stackblitz.com/img/open_in_stackblitz.svg)](https://stackblitz.com/github/Forevka/canvas-word/tree/main/examples/embed-multi) |
| [`ribbon-customization`](./ribbon-customization) | Customize the ribbon (`customizeRibbon`) — reorder/remove built-ins, add tabs/buttons, macros + popups | None (static HTML) | [![Open in StackBlitz](https://developer.stackblitz.com/img/open_in_stackblitz.svg)](https://stackblitz.com/github/Forevka/canvas-word/tree/main/examples/ribbon-customization) |
| [`command-registry`](./command-registry) | Register custom commands + keyboard shortcuts (`commands`) — cross-platform `Mod+…` bindings, `handle.runCommand`, built-ins always win | None (static HTML) | [![Open in StackBlitz](https://developer.stackblitz.com/img/open_in_stackblitz.svg)](https://stackblitz.com/github/Forevka/canvas-word/tree/main/examples/command-registry) |
| [`custom-fonts`](./custom-fonts) | Supply your own fonts (`fonts`) — a self-hosted PT Serif (4 faces + sizing), hide a built-in, export with the font embedded | None (static HTML) | [![Open in StackBlitz](https://developer.stackblitz.com/img/open_in_stackblitz.svg)](https://stackblitz.com/github/Forevka/canvas-word/tree/main/examples/custom-fonts) |
| [`editor-constructor`](./editor-constructor) | Interactive config builder — toggle every constructor option, live preview, copy a `new WordCanvas({...})` snippet | Vite | [![Open in StackBlitz](https://developer.stackblitz.com/img/open_in_stackblitz.svg)](https://stackblitz.com/github/Forevka/canvas-word/tree/main/examples/editor-constructor) |

Run locally from the repo root (see each folder's README). Most import the built
`dist-lib` bundle, so build the library first: `npm run build:lib`.

## Open in StackBlitz

Each **Open in StackBlitz** button boots the example in the browser against the
**published** `@forevka/wordcanvas` from npm. Caveats:

- The **Vite** examples (`embed-live`, `playground`, `query-edit`, `merge-docs`,
  `embed-multi`, `editor-constructor`) install the package and run under Vite.
- The **static** examples (`embed-offline`, `ribbon-customization`,
  `command-registry`, `custom-fonts`)
  use an import map pointing at a local `dist-lib` and are primarily meant to run
  locally; on StackBlitz they may need a package.json / CDN import-map tweak.
- `query-edit` and `merge-docs` need the **`./query` subpath** (the latter also
  uses `mergeDocuments` / `mergeAll`), which must be present in the published
  version (`npm view @forevka/wordcanvas exports`) — publish a current version if
  your registry copy predates the query/edit release.
