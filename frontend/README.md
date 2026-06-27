# @forevka/wordcanvas

An embeddable, canvas-rendered, Word-compatible document editor. It paints pages
to a `<canvas>` (no `contenteditable`), paginates exactly like Word, imports and
exports real `.docx` and `.pdf`, and — when pointed at a backend — supports live
multi-user collaboration with presence.

The published bundle is **self-contained**: it has **zero runtime dependencies**
(the layout engine, font tooling, and DOCX/PDF pipelines are all inlined and
code-split). You can drop it onto a page with a plain `<script type="module">` or
import it from any bundler. The same package also runs **headless on Node** for
server-side `.docx`/`.pdf` generation — see [Server-side (Node)](#server-side-node).

**▶ Live demo:** [doc-editor.forevka.dev](https://doc-editor.forevka.dev/) (full
feature showcase) · [minimal offline embed](https://doc-editor.forevka.dev/examples/offline/).

## How it compares

If you have shopped for an embeddable Word editor, you have met Syncfusion
Document Editor, OnlyOffice, and DevExpress Rich Text Editor. They are mature
and cover more of the Word long tail than this package does. They also ask for a
commercial per-seat license, and two of the three want a server running before
the editor renders a page.

WordCanvas takes the other trade. It ships under MIT, paints to a `<canvas>`
the way Google Docs has since 2021, and runs fully in the browser with zero
runtime dependencies. You only stand up a backend if you want live
collaboration; reading, editing, and DOCX/PDF export all work offline.

| | WordCanvas | Syncfusion Document Editor | OnlyOffice | DevExpress Rich Text |
|---|---|---|---|---|
| License | MIT | Commercial seat | AGPL or commercial | Commercial seat |
| Rendering | Canvas | DOM | Canvas (in an iframe) | DOM |
| Server required to render | No | For some file conversions | Yes (Document Server) | Yes (.NET backend) |
| Runtime dependencies | Zero | Several | Bundled suite | .NET stack |
| DOCX import + export | In-browser | Yes | Yes | Yes |
| PDF export | Page-accurate, in-browser | Yes | Yes | Yes |
| Nested content controls | **Any depth** — incl. fields inside | Flat only | Flat only | Flat only |
| Word fields (PAGE/DATE/IF…) | Editable objects, round-trip | Yes | Yes | Yes |
| Live collaboration | Built in (opt-in backend) | Add-on | Built in | Add-on |
| Primary target | Any JS app | Angular/React/Vue | Iframe / full suite | Blazor / .NET |

Where WordCanvas goes *past* them: **nested content controls** — a `w:sdt`
inside a `w:sdt`, to any depth, with complex Word fields nested inside —
round-trip through `.docx` without flattening. That matters for C#/OOXML
generated reports, whose section→field control structure the commercial editors
collapse on edit.

Right-to-left, bidirectional, and CJK editing now work out of the box (full
UAX#9 bidi with multi-run caret/selection, RTL `.docx` round-trip, and CJK
kinsoku line-breaking). **Mathematical equations** are supported too: Presentation
MathML is the canonical form, typeset with the bundled STIX Two Math font (real
math glyphs, growing delimiters, display-size big operators) and round-tripped to
`.docx` as OMML (inline `m:oMath` + display `m:oMathPara`); the visual equation
editor accepts MathML or LaTeX, with a live preview, and equations are
selectable/right-click-editable like images. Where the commercial editors still
win today: charts, a CJK font bundled in the box (CJK renders on-screen via system
fonts but needs a registered font to embed in PDF/DOCX export — see
`cjk.fallbackFont`), and an enterprise support contract. The full breakdown,
including what each one does better, lives in
[Best embeddable JS Word editors](https://forevka.dev/articles/best-embeddable-js-word-editors/).

## Install

```sh
npm install @forevka/wordcanvas
```

## Quick start (offline)

Omit `backendUrl` and the editor runs fully local — no network, no sync, no
share. Perfect for a standalone document editor.

```ts
import { WordCanvas } from "@forevka/wordcanvas";

const editor = new WordCanvas({
  container: document.getElementById("editor")!,
});

editor.on("ready", () => console.log("editor mounted"));
```

```html
<div id="editor" style="position:fixed; inset:0;"></div>
```

### No build step

The bundle is a standalone ES module, so an import map is all you need:

```html
<div id="editor" style="position:fixed; inset:0;"></div>
<script type="importmap">
  { "imports": { "@forevka/wordcanvas": "/node_modules/@forevka/wordcanvas/dist-lib/wordcanvas.js" } }
</script>
<script type="module">
  import { WordCanvas } from "@forevka/wordcanvas";
  new WordCanvas({ container: document.getElementById("editor") });
</script>
```

## Online + collaboration

Pass `backendUrl` to turn on sync. Opening a document then auto-publishes it and
exposes a shareable link; edits sync live and presence events fire. The embedder
owns identity — pass a `user` so carets and edits are attributed.

```ts
import { WordCanvas } from "@forevka/wordcanvas";

const params = new URLSearchParams(location.search);

const editor = new WordCanvas({
  container: document.getElementById("editor")!,
  backendUrl: "https://api.example.com",
  user: { id: "u-42", firstName: "Ada", lastName: "Lovelace" },
  // Join an existing session if the share link carried one (the editor's own
  // share links use ?collab=<docId>):
  docId: params.get("collab") ?? undefined,
});

editor.on("shared", ({ url }) => navigator.clipboard.writeText(url));
editor.on("presence", ({ participants }) => renderAvatars(participants));
```

A runnable version of this lives in [`examples/embed-live`](../examples/embed-live).
A minimal offline embed lives in [`examples/embed-offline`](../examples/embed-offline).

## API

```ts
new WordCanvas(options)
```

| Option        | Type                                   | Notes                                                                 |
| ------------- | -------------------------------------- | --------------------------------------------------------------------- |
| `container`   | `HTMLElement`                          | **Required.** Element to mount into.                                  |
| `backendUrl`  | `string`                               | Online iff provided. Omit for a fully offline editor.                 |
| `docId`       | `string`                               | Open this document on load (online only).                             |
| `user`        | `{ id, firstName, lastName }`          | Identity for attribution + presence.                                  |
| `onShareLink` | `(url, docId) => void`                 | Override how the share link is surfaced (default: built-in dialog).   |
| `readonly`    | `boolean`                              | Mount as a view-only viewer (see below). Default `false`.            |

Methods: `whenReady(): Promise<EditorHandle>`, `openDocx(file)`, `share()`,
`getDocId()`, `getShareLink()`, `destroy()`, and `on(event, handler)` /
`off(event, handler)`.

Events: `ready`, `shared`, `userEntered`, `userLeave`, `presence`.

```ts
// Load a .docx the user picked (auto-publishes + shares when online):
input.addEventListener("change", async () => {
  await editor.openDocx(input.files![0]);
});
```

### Read-only / viewer mode

Pass `readonly: true` to mount a view-only viewer:

```ts
const viewer = new WordCanvas({ container, readonly: true });
await viewer.openDocx(file); // or viewer.setDocument(doc)
```

The document still renders, scrolls, and stays selectable and copyable (Ctrl+C,
right-click → Copy), and `Ctrl+F` find works. The editing chrome is hidden (no
ribbon or ruler; the find bar drops Replace) and **every mutation is a no-op** —
typing, paste, undo/redo, drag-resize, and the programmatic editing paths all
short-circuit. In an online session a read-only client still joins the
collaboration and receives live remote edits; it just can't author them.

## Server-side (Node)

The same package runs **headless on Node** — build or import a document, then
export it to `.docx`/`.pdf` with no browser and no DOM. This is the same pipeline
the editor uses, so a server-rendered file paginates identically to the on-screen
document. The pipeline subpaths are isomorphic via [conditional exports](https://nodejs.org/api/packages.html#conditional-exports):
the `node` condition resolves a self-contained Node build (bundled clone fonts +
PDF engine, still zero runtime dependencies), and a browser bundler gets the
browser build automatically — the same `import` works in both.

Compose a document with the fluent [builder](../BUILDER.md), install the
measurement host once (it loads the bundled fonts), then export:

```ts
import { writeFile } from "node:fs/promises";
import { DocumentBuilder } from "@forevka/wordcanvas/builder";
import { installMeasureHost } from "@forevka/wordcanvas/export/measure";
import { runExport } from "@forevka/wordcanvas/export";

const doc = DocumentBuilder.create()
  .paragraph("Quarterly report").withStyle("Heading1")
  .paragraph("Generated on the server — no browser, no DOM.")
  .build();

await installMeasureHost();                    // once per process, before exporting
const { bytes } = await runExport(doc, "pdf"); // or "docx"
await writeFile("report.pdf", bytes);
```

Parse an uploaded `.docx` back into the editable model:

```ts
import { runImport } from "@forevka/wordcanvas/import";

const { doc, warnings } = runImport(new Uint8Array(uploadedBytes));
```

| Subpath                              | Purpose                                                        |
| ------------------------------------ | ------------------------------------------------------------- |
| `@forevka/wordcanvas/builder`        | Fluent, programmatic document composer (see [BUILDER.md](../BUILDER.md)). |
| `@forevka/wordcanvas/import`         | `.docx` → document model.                                     |
| `@forevka/wordcanvas/export`         | Document model → `.docx`/`.pdf` bytes.                        |
| `@forevka/wordcanvas/export/measure` | Install the headless font/measurement host (await once first). |
| `@forevka/wordcanvas/recalc-docx`    | Patch a `.docx`'s cached TOC page numbers in place.           |
| `@forevka/wordcanvas/generate-toc`   | Generate a Word TOC field result headless.                   |

## Notes & limitations

- **Multiple instances per page are supported** — the chrome is class-scoped under
  a per-instance root and each instance mounts independently, so you can run several
  editors at once. `destroy()` removes an instance's chrome and its floating panels.
  See the [`embed-multi`](../examples/embed-multi) example.
- The editor injects its own stylesheet and mounts its full chrome (ribbon, ruler,
  outline, status bar) inside the container; give it a sized element.
- The bundle is large on disk (~20 MB unminified, code-split) because it ships a
  full layout engine plus DOCX/PDF tooling. The initial entry chunk is small; the
  heavy editor app and the export/import workers are lazily loaded on demand.

## License

MIT © Bohdan Lushchyk
