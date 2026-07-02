# Query & Edit API

[![Open in StackBlitz](https://developer.stackblitz.com/img/open_in_stackblitz.svg)](https://stackblitz.com/github/Forevka/canvas-word/tree/main/examples/query-edit)

Worked example of the headless **query + edit** API — `@forevka/wordcanvas/query`
— the isomorphic core you'd use in Node, a worker, or the C# bindings (the rough
analog of .NET's `WordprocessingDocument`). It builds a document with the
`DocumentBuilder`, shows it in a live editor on the right, and drives every query
and edit from the panel on the left.

## What it shows

- **Query (read-only)** over a plain `Document`:
  - `getParagraphs` / `getSections` / `getStyles`
  - `findParagraphs("lorem")` — find by substring/RegExp
  - `getSdts` + `sdtText` — content controls and their text
  - `positionOfText` + `rangeText` — offset-free addressing
  - `getPages` — "what's on page N" (runs a layout pass)
- **Edit** via `DocumentEditor` (rides the operation engine, so undo/redo is free):
  - `setSdtText` — fill a content control
  - `replaceAllText` — find/replace across the doc as one undoable step
  - `moveBlock` — reorder a block
  - `undo` / `redo`
  - Each edit is pushed into the live editor with `WordCanvas.setDocument`, so you
    watch it land.

## Run (locally, from the repo root)

```sh
npm run dev --workspace @cw/example-query-edit   # dev server on http://localhost:5182
```

In dev, the `/builder` and `/query` subpaths are aliased to the frontend **source**
(see `vite.config.ts`), so it runs against the current tree without a lib rebuild.

Production build (requires the library bundle):

```sh
npm run build:lib                                    # build @forevka/wordcanvas
npm run build --workspace @cw/example-query-edit     # bundle + copy dist-lib
```

## StackBlitz note

The **Open in StackBlitz** button above installs `@forevka/wordcanvas` from npm.
The query/edit surface lives on the `./query` subpath — make sure a **published**
version exposes it (`npm view @forevka/wordcanvas exports`). If your registry copy
predates the query/edit release, publish a newer version first; the `./import` and
`./export` subpaths and the `.`/`./builder` entries are already published.
