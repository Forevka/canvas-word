# Examples

Standalone examples of embedding [`@forevka/wordcanvas`](../frontend) as a library,
each consuming the package by name the way a third-party integrator would.

| Example | What it shows | Tooling |
| --- | --- | --- |
| [`embed-offline`](./embed-offline) | Smallest possible integration — fully local editor (no backend), via a browser import map | None (static HTML) |
| [`embed-live`](./embed-live) | Online editing + live collaboration + presence + share links (migrated from the old `/live` page) | Vite |

Both require the library to be built first (`npm run build:lib` from the repo
root), since they import the built `dist-lib` bundle. See each folder's README for
run instructions.
