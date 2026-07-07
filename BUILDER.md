# Document Builder — programmatic composition + live preview

`@forevka/wordcanvas/builder` is a fluent API for generating documents in
JS/TS, the use case that usually forces a C#/Java backend with an OOXML SDK.
Describe the document in code, bind a JSON data model, preview it live in the
embedded editor, and export DOCX/PDF. The same code runs in the browser and in
Node.

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
returns a deep-cloned `Document`: the same object the editor renders, the
collaboration layer replicates, and the exporters write.

It lives at `frontend/src/builder/` (published as the `./builder` subpath)
because `fromTemplate` reuses the docx import pipeline, which is pure, DOM-free
TypeScript. So the whole builder works in Node for server-side generation,
following the same precedent as the `./import` subpath the backend consumes.

### Chaining semantics

`paragraph()` eagerly appends a paragraph and returns a paragraph scope. The
scope's styling methods mutate that paragraph; any block-starting call
(`paragraph`, `table`, `image`, `list`, …) delegates back to the parent scope,
so the chain "pops" automatically with no seal/end step (`end()` exists for
explicitness but is never required).

```ts
b.paragraph("Title").withStyle("Heading1")   // paragraph scope
 .paragraph("Body")                           // pops, starts the next one
 .table([["a", "b"]])                         // pops to the document scope
```

Character formatting in a paragraph scope (`bold()`, `color()`, `font()`,
`fontSize()`, `link()`, …) patches **every run already in the paragraph** and
becomes the default for runs added later by `text()`. "Make this paragraph
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
(`builder.warnings`), not an error. Register new styles with
`b.style({ id, name, basedOn, char, para })` **before** applying them:
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
so the document is portable across the editor, browser workers, and Node.
Import warnings surface on `builder.warnings`.

### Data binding: declarative rebuild

There is no merge-field/placeholder engine: the *code is the template*. Write a
function `(data) => Document` and re-run it when data changes.
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
.table(rows2d | callback, { colFractions, headerRow, style })  // style = a preset name
.image(url | { data, mime }, { widthPx, heightPx, align, wrap })  // dims required
.list(items, { kind, listId, level }) / .bulletList(items) / .numberedList(items)
.pageBreak()                         // next block starts a new page
.columnBreak()                       // next block starts a new newspaper column
.bookmarkRange(name, s => …)         // bookmark every paragraph the callback adds

// paragraph scope (plus all block-scope methods, which pop back):
.withStyle(id) .text(t, patch?) .bold() .italic() .underline() .strikethrough()
.color(c) .highlight(c) .fontSize(px) .font(f) .link(url)
.superscript() .subscript() .letterSpacing(px) .hidden()
.align(a) .spacing({ before, after, lineHeight }) .indent({ left, right, firstLine })
.keepWithNext() .end()
// inline fields (tagged result run + registered FieldDef):
.field(spec) .pageField(numFmt?) .numPagesField(numFmt?) .dateField(fmt?) .timeField(fmt?)
.ifField(a, op, b, ifTrue, ifFalse) .customField(instr, resultText, { name? })
.crossReference(bookmarkName, { kind: "ref" | "pageRef" })
// inline content controls (SDT):
.contentControl(props, text) .richTextControl(t) .plainTextControl(t) .checkbox(checked)
.dropDown(sel, items) .comboBox(sel, items) .dateControl(t, fmt?)
// inline notes + bookmarks:
.footnote(text | (s => …))           // auto-numbered marker + note body
.bookmark(name, text)                // bookmark exactly the appended run

// document scope only:
.header(s => …, { variant: "default" | "first" | "even" }) .footer(…)
.pageSetup({ pageSize, orientation, margins, columns, headerDistancePx, footerDistancePx, pageNumberStart })
.tableOfContents({ maxLevel?, hyperlink?, title?, leader?, levels? })  // resolved at build()
.sectionBreak({ pageSize?, orientation?, margins?, columns?, pageNumberStart?, header?, footer?, … })
.style(namedStyle)            // register a named style (before use)
.defaultStyle(id)             // set + inherit the document default style (call early)
.listDefinition(id, spec)     // custom bullet/number/multilevel/levels list
.tableStylePreset(name, preset)  // reusable cell formatting for .table({ style })
.warnings                 // readonly BuilderWarning[]
.build()                  // → Document (deep clone; builder stays usable)
```

Unit helpers: `inches(n)`, `cm(n)`, `pt(n)`, `twips(n)` (all → px @96dpi),
`PAGE_SIZES` (Letter, Legal, A4, A3, Tabloid), `bytesToDataUrl(bytes, mime)`.

### Computed fields

Built-in fields are first-class objects (a typed `FieldSpec` + a registered
`FieldDef`), authored inline in a paragraph:

```ts
b.paragraph("Page ").pageField().text(" of ").numPagesField()    // {page} / {pages} tokens
 .paragraph("Today is ").dateField("MMMM d, yyyy")                // materialized now
 .paragraph().ifField("2", ">", "1", "in stock", "back-ordered") // chosen branch
```

`pageField`/`numPagesField` emit a live `{page}`/`{pages}` token the layout engine
re-resolves per page (in the body **and** in bands);
`dateField`/`timeField`/`ifField` materialize their result once (pass `{ now }` for
a deterministic date in tests). `customField(instruction, resultText)` is the escape
hatch for any other field (`SEQ`, `STYLEREF`, …); `crossReference(name)` is sugar over
it for `REF`/`PAGEREF` to a bookmark. IF branch results are plain text.

### Content controls (SDT)

```ts
b.paragraph("Choose: ").dropDown("One", [{ display: "One", value: "1" }], { alias: "Choice" })
 .paragraph("Agree ").checkbox(true)
```

`contentControl(props, text)` takes raw `SdtProps`; `richTextControl`/`plainTextControl`/
`checkbox`/`dropDown`/`comboBox`/`dateControl` are per-kind sugar. Authoring only —
in-place data rebinding that preserves user edits across rebuilds is out of scope.

### Footnotes & bookmarks

`footnote(text | callback)` appends an auto-numbered marker and registers the note body
(a string is one paragraph; a `StoryBuilder` callback is rich/multi-paragraph). `bookmark(name, text)`
bookmarks exactly the appended run; `bookmarkRange(name, callback)` bookmarks every paragraph the
callback adds (a multi-block span). Footnote numbers follow insertion order — use a fresh builder
per data-driven rebuild (the documented rebuild model) so numbers don't accumulate.

### Table of contents

`tableOfContents(opts)` is **deferred**: it drops a placeholder anchor and sets the doc's
`TOC` field instruction now, then `build()` generates the entries from the document's
headings, so headings added *after* the call are still included. Mark headings with
`.withStyle("Heading1")` (or an outline level). The entries are live (page numbers resolve
at layout) and round-trip to `.docx` as a real `TOC` field.

### Sections & columns

`sectionBreak(opts)` ends the current section and starts a new one on the next page, with its
own page size/orientation/margins, **newspaper columns** (`columns: { count, gapPx? }`),
page-number restart, and header/footer band callbacks. `columnBreak()` starts the next
newspaper column (like `pageBreak()`, it applies to the following block).

### Custom styles, list definitions & table-style presets

`style({ id, name, basedOn, char, para })` registers a named style (apply with `.withStyle(id)`);
`defaultStyle(id)` makes one the document default so later content inherits it (call early).
`listDefinition(id, spec)` registers a custom list (`{ kind: "bullet", char }`,
`{ kind: "number", format, suffix? }`, `{ kind: "multilevel" }`, or raw `{ levels }`),
referenced via `.list(items, { listId: id })`. `tableStylePreset(name, preset)` defines reusable
cell formatting (header styling, borders, shading, zebra striping) applied with
`.table(rows, { style: name })` — **builder-only sugar**: it resolves to concrete cell
properties at build time (explicit per-cell `CellSpec` values win), so nothing style-id-like
enters the model. Built-in presets: `plain`, `grid`, `headerBand`, `striped`.

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
fonts on disk, zero runtime deps) for server/headless rendering; the `default`
condition resolves a browser variant. The same import works in both environments;
browser embedders can still export via the editor toolbar or `exportDocument`.

## Testing

- `frontend/src/builder/builder.test.ts` — model-shape assertions, style
  resolution/precedence, scope semantics, clone isolation, deterministic ids.
- `frontend/src/builder/builderFeatures.test.ts` — the field/SDT/footnote/bookmark/
  TOC/section/list/preset surface (per-feature model-shape assertions; `now` injected).
- `frontend/src/builder/templateRoundtrip.test.ts` — builder → export →
  `fromTemplate` → compose → export → re-import; the exporter manufactures the
  template fixture, so no binary is committed.
- `frontend/src/builder/featuresRoundtrip.test.ts` — a doc exercising the new
  features survives export → re-import (fields/SDT/footnotes/bookmarks/TOC persist).

## Not supported (yet)

Out of scope. The model supports several of these already; the builder just
doesn't author them:

- Content-control (SDT) **data rebinding** — controls are authored, but in-place
  partial rebinding that preserves user edits across data changes is not.
- Nested tables via the fluent API (possible by composing cells, not first-class).
- Floating anchored images (only `block` / `square` wrap).
- `{{placeholder}}` merge syntax inside templates (declarative rebuild is the
  binding model).
- Charts, shapes, text boxes; tracked changes; comments.
- Image natural-size auto-measure (explicit `widthPx`/`heightPx` required).
