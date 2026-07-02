# Merge documents

[![Open in StackBlitz](https://developer.stackblitz.com/img/open_in_stackblitz.svg)](https://stackblitz.com/github/Forevka/canvas-word/tree/main/examples/merge-docs)

Worked example of the headless **merge / append** API — the drop-in replacement
for Syncfusion `ImportContent` + hand-rolled `ImportStyles`. It builds a few
independently-authored **parts** with the `DocumentBuilder`, folds them into one
report with `@forevka/wordcanvas/query`, shows the result in a live editor on the
right, and applies a per-section **content footer** (a table with a logo).

## What it shows

- **`DocumentEditor.append(source, opts)`** — append a part after the current
  document as one undoable step; every id space (styles, lists, content controls,
  fields, notes, bookmarks) is reconciled so nothing collides. The parts here share
  an id seed and all define an `intro` bookmark on purpose, so you can watch the
  clashes get renamed (see **warnings** / **bookmarks remapped** in the report).
- **`mergeAll([...], opts)`** — fold N documents in one call.
- **Section seam** — `sectionBreak: "nextPage"` starts each part on a new page and
  keeps its own page geometry (the parts use Letter / A4 / Legal so the boundaries
  are visible).
- **`setSectionFooter(index, blocks)`** — apply a branded footer to a section after
  merging: a borderless `[ logo | right-aligned address ]` table over a centered
  `Page X of Y`. The page number is a live field (no placeholder), and the address
  wraps on its own (no manual line-splitting). The footer `Block[]` is authored with
  the same `DocumentBuilder` band surface (`.footer(f => f.table(...).paragraph(...))`).

## Run it

```sh
npm install          # once, from the repo root
npm run dev:merge-docs
```

Then open the printed local URL. Click **Append Body**, **Append Appendix**, then
**Add footer to last section** — or **Merge all at once** — and use **Undo/Redo**.

## Note on StackBlitz / npm

This example imports `@forevka/wordcanvas/query`. Until a package version that
exposes the `./query` subpath (including `mergeDocuments` / `mergeAll`) is
published, the StackBlitz button and a plain `npm install` will only resolve the
`.` / `/builder` / `/import` / `/export` entries. Running from the monorepo (the
`dev:merge-docs` script) always works — the dev server aliases `/builder` + `/query`
to the frontend source.
