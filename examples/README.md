# Examples

Standalone examples of embedding [`@forevka/wordcanvas`](../frontend) as a library,
each consuming the package by name the way a third-party integrator would.

| Example | What it shows | Tooling |
| --- | --- | --- |
| [`embed-offline`](./embed-offline) | Smallest possible integration — fully local editor (no backend), via a browser import map | None (static HTML) |
| [`embed-live`](./embed-live) | Online editing + live collaboration + presence + share links (migrated from the old `/live` page) | Vite |
| [`playground`](./playground) | Programmatic `DocumentBuilder` playground — edit code, see the document rebuild live | Vite |
| [`embed-multi`](./embed-multi) | Multiple independent editors on one page; add/remove instances at runtime | Vite |
| [`ribbon-customization`](./ribbon-customization) | Customize the ribbon (`customizeRibbon`) — reorder/remove built-ins, add tabs/buttons, macros + popups | None (static HTML) |
| [`editor-constructor`](./editor-constructor) | Interactive config builder — toggle every constructor option, live preview, copy a `new WordCanvas({...})` snippet | Vite |

They require the library to be built first (`npm run build:lib` from the repo
root), since they import the built `dist-lib` bundle. See each folder's README for
run instructions.
