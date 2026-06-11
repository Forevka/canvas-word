# Document builder playground

Interactive playground for the `@forevka/wordcanvas/builder` fluent API: write
builder code on the left, feed it a JSON data model, and watch the document
render live in the editor on the right.

- **Code pane** — the body of `async (data, ctx) => Document`. `data` is your
  parsed JSON; `ctx` provides `DocumentBuilder`, the uploaded `template`
  (ArrayBuffer or `null`), the unit helpers (`inches`, `cm`, `pt`, `twips`,
  `PAGE_SIZES`), and `bytesToDataUrl`. End with `return b.build()`.
- **Data pane** — the JSON data model. Every edit re-runs your code with the
  new data and swaps the result into the editor (declarative rebuild).
- **Load template…** — pick a `.docx`; its named styles, list definitions,
  page setup, and header/footer bands become the base your code composes
  against (`DocumentBuilder.fromTemplate(ctx.template)`).
- The preview is a full editor: the generated document is editable, and the
  toolbar exports DOCX/PDF. A rebuild (code/data change) replaces the document
  and discards manual edits.

Both panes persist to `localStorage`; the template does not (re-upload after a
reload).

## Run

From the repo root:

```sh
npm run dev:playground     # dev server on http://localhost:5181
```

Production build (requires the library bundle):

```sh
npm run build:lib                                  # build @forevka/wordcanvas
npm run build --workspace @cw/example-playground   # bundle + copy dist-lib
```

The build keeps `@forevka/wordcanvas` (and its `/builder` subpath) external and
resolves them at runtime through the import map in `index.html`, pointing at
the copied `./wordcanvas/` bundle — the same pattern as `examples/embed-live`.
