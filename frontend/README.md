# @forevka/wordcanvas

An embeddable, canvas-rendered, Word-compatible document editor. It paints pages
to a `<canvas>` (no `contenteditable`), paginates exactly like Word, imports and
exports real `.docx` and `.pdf`, and — when pointed at a backend — supports live
multi-user collaboration with presence.

The published bundle is **self-contained**: it has **zero runtime dependencies**
(the layout engine, font tooling, and DOCX/PDF pipelines are all inlined and
code-split). You can drop it onto a page with a plain `<script type="module">` or
import it from any bundler.

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

## Notes & limitations

- **One instance per page** in this version — the chrome uses id-based CSS and the
  app bootstraps once on import. Constructing a second `WordCanvas` throws.
- The editor injects its own stylesheet and mounts its full chrome (ribbon, ruler,
  outline, status bar) inside the container; give it a sized element.
- The bundle is large on disk (~20 MB unminified, code-split) because it ships a
  full layout engine plus DOCX/PDF tooling. The initial entry chunk is small; the
  heavy editor app and the export/import workers are lazily loaded on demand.

## License

MIT © Bohdan Lushchyk
