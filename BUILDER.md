# Document Builder — programmatic composition + live preview

`@forevka/wordcanvas/builder` is a fluent API for generating documents in
JS/TS — the use case that traditionally forces a C#/Java backend with an OOXML
SDK. Describe the document in code, bind a JSON data model, preview it live in
the embedded editor, and export DOCX/PDF. The same code runs in the browser and
in Node.

```ts
import { DocumentBuilder } from "@forevka/wordcanvas/builder";

const build = async (data) => {
  const b = await DocumentBuilder.fromTemplate(templateBytes); // styles from a .docx
  return b
    .paragraph(data.title).withStyle("Heading1")
    .paragraph(`Prepared for ${data.customer} on ${data.date}`).italic()
    .table([
      ["Item", "Qty"],
      ...data.items.map((i) => [i.name, String(i.qty)]),
    ], { headerRow: true })
    .bulletList(data.notes)
    .footer((f) => f.paragraph("Page {page} of {pages}").align("center"))
    .build();
};

// Live preview: rebuild on every data change and swap into the editor.
const wc = new WordCanvas({ container });
await wc.setDocument(await build(data));
```

Try it interactively: `npm run dev:playground` (see `examples/playground`).

## Design

The editor's document model is plain data (`Document` → `Block[]` → `Paragraph`
→ `Run[]`, see `shared/src/model/document.ts`), so the builder is a thin fluent
layer that mints model objects — no editor coupling, no hidden state. `build()`
returns a deep-cloned `Document`; the same object the editor renders, the
collaboration layer replicates, and the exporters write.

It lives at `frontend/src/builder/` (published as the `./builder` subpath)
because `fromTemplate` reuses the docx import pipeline — which is pure,
DOM-free TypeScript, so the whole builder works in Node for server-side
generation, following the same precedent as the `./import` subpath the backend
consumes.

### Chaining semantics

`paragraph()` eagerly appends a paragraph and returns a paragraph scope. The
scope's styling methods mutate that paragraph; any block-starting call
(`paragraph`, `table`, `image`, `list`, …) delegates back to the parent scope —
so the chain "pops" automatically with no seal/end step (`end()` exists for
explicitness but is never required).

```ts
b.paragraph("Title").withStyle("Heading1")   // paragraph scope
 .paragraph("Body")                           // pops, starts the next one
 .table([["a", "b"]])                         // pops to the document scope
```

Character formatting in a paragraph scope (`bold()`, `color()`, `font()`,
`fontSize()`, `link()`, …) patches **every run already in the paragraph** and
becomes the default for runs added later by `text()` — "make this paragraph
bold" is the dominant authoring intent. Mixed formatting uses explicit run
patches: `.paragraph("see ").text("docs", { link: url })`.

Structures that nest (tables, header/footer bands, table cells) use callback
scopes instead of the flat chain:

```ts
b.table((t) => t
  .row((r) => r.cell("Merged", { colSpan: 2, shading: "#eef" }))
  .row((r) => r.cell("a").cell((cell) => cell.paragraph("rich").bulletList(["x"]))),
  { colFractions: [1, 2] })
```

Cells in the data-driven shape accept `string` or a `CellSpec`
(`{ text, colSpan, rowSpan, shading, borders, margin, style, align }`).

### Named styles

`withStyle(id)` works like applying a style in the editor: it records the
reference (`para.style.namedStyle`) **and** patches exactly the fields the
style defines (resolved through the `basedOn` chain) onto the paragraph and its
runs. Call order is precedence — direct formatting applied *after* `withStyle`
wins; fields the style doesn't define survive. Unknown style ids are a warning
(`builder.warnings`), not an error, and register new styles with
`b.style({ id, name, basedOn, char, para })` **before** applying them —
resolution happens at call time, not at `build()`.

### Templates

`DocumentBuilder.fromTemplate(docxBytes)` imports the .docx and keeps what a
template is *for*:

| Kept | Discarded (default) |
|---|---|
| stylesheet (named styles, `basedOn` chains) | body content (`keepBody: true` to append after it) |
| list definitions | footnotes (body-anchored) |
| page setup (size, margins, columns, band distances) | bookmarks/content-controls whose anchors lived in the body |
| header/footer bands, incl. first/even variants | |

Embedded images in kept stories are inlined as `data:` URLs (never `blob:`),
so the resulting document is portable across the editor, browser workers, and
Node. Import warnings surface on `builder.warnings`.

### Data binding: declarative rebuild

There is no merge-field/placeholder engine: the *code is the template*. Write a
function `(data) => Document` and re-run it when data changes —
`WordCanvas.setDocument` swaps the result into the live editor, preserving zoom
and scroll position so the preview is stable. Rebuilds discard the undo stack
and any manual edits (by design — the document is a projection of the data),
and, like `openDocx`, fork away from any live collaboration session.

The only computed fields are the `{page}` / `{pages}` tokens in header/footer
text (substituted per page at layout time; `{page:roman|Roman|alpha|Alpha}`
formats supported).

### IDs

Each builder mints block/cell ids from its own namespace (`bldXXXX-N` via
`shared/src/ids.ts`), disjoint from template-derived and editor-session ids, so
generated content can never alias collaborating clients' blocks. Pass
`idSeed` in `create`/`fromTemplate` options for deterministic ids in tests.

## API summary

```ts
DocumentBuilder.create(opts?: { pageSize, margins, stylesheet, idSeed })
DocumentBuilder.fromTemplate(docx, opts?: { keepBody, idSeed })  // async (the only await)

// block scope (document body, bands, cells):
.paragraph(text?, charPatch?)        → paragraph scope
.table(rows2d | callback, { colFractions, headerRow })
.image(url | { data, mime }, { widthPx, heightPx, align, wrap })  // dims required
.list(items, { kind, listId, level }) / .bulletList(items) / .numberedList(items)
.pageBreak()                         // next block starts a new page

// paragraph scope (plus all block-scope methods, which pop back):
.withStyle(id) .text(t, patch?) .bold() .italic() .underline() .strikethrough()
.color(c) .highlight(c) .fontSize(px) .font(f) .link(url)
.align(a) .spacing({ before, after, lineHeight }) .indent({ left, right, firstLine })
.keepWithNext() .end()

// document scope only:
.header(s => …, { variant: "default" | "first" | "even" }) .footer(…)
.pageSetup({ pageSize, orientation, margins, columns, headerDistancePx, footerDistancePx, pageNumberStart })
.style(namedStyle)        // register before use
.warnings                 // readonly BuilderWarning[]
.build()                  // → Document (deep clone; builder stays usable)
```

Unit helpers: `inches(n)`, `cm(n)`, `pt(n)`, `twips(n)` (all → px @96dpi),
`PAGE_SIZES` (Letter, Legal, A4, A3, Tabloid), `bytesToDataUrl(bytes, mime)`.

## Server-side (Node)

The published `./builder` entry is editor-free and runs in Node as-is:

```ts
import { DocumentBuilder } from "@forevka/wordcanvas/builder";
const doc = DocumentBuilder.create().paragraph("hi").build();
```

Pair it with the export pipeline to produce binaries. On Node you must install
the headless measurement host once (it loads the bundled clone fonts) before the
first export/layout:

```ts
import { installMeasureHost } from "@forevka/wordcanvas/export/measure";
import { runExport } from "@forevka/wordcanvas/export";
await installMeasureHost();
const { bytes } = await runExport(doc, "docx", {});             // or "pdf"
```

These pipeline subpaths (`./import`, `./export`, `./export/measure`,
`./recalc-docx`, `./generate-toc`) are published as **conditional exports**: the
`node` condition resolves a self-contained, Node-targeted bundle (bundled pdfkit +
fonts on disk, zero runtime deps) for server/headless rendering, and the `default`
condition resolves a browser variant. The same import works in both environments;
browser embedders can still export via the editor toolbar or `exportDocument`.

## Testing

- `frontend/src/builder/builder.test.ts` — model-shape assertions, style
  resolution/precedence, scope semantics, clone isolation, deterministic ids.
- `frontend/src/builder/templateRoundtrip.test.ts` — builder → export →
  `fromTemplate` → compose → export → re-import; the exporter manufactures the
  template fixture, so no binary is committed.

## Not supported (yet)

Intentionally out of the first version — the model supports several of these
already; the builder just doesn't author them:

- Footnotes/endnotes authoring
- Bookmarks and cross-references
- TOC generation
- Content-control (SDT) data binding / in-place partial rebinding that
  preserves user edits across data changes
- Section breaks / multi-section page setup (one section per document)
- Newspaper columns per-block (whole-document `pageSetup.columns` only)
- Computed fields in the body — PAGE/NUMPAGES/dates/formulas (`{page}`/`{pages}`
  band tokens are the only computed text)
- Table styles (direct cell formatting only) and nested tables via the fluent
  API (possible by composing cells, not first-class)
- Floating anchored images (only `block` / `square` wrap)
- `{{placeholder}}` merge syntax inside templates (declarative rebuild is the
  binding model)
- Charts, shapes, text boxes; tracked changes; comments
- Image natural-size auto-measure (explicit `widthPx`/`heightPx` required)
