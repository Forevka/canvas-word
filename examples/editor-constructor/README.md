# Editor constructor / config builder

An interactive builder for the **`new WordCanvas({…})`** constructor. Toggle any
option in the form on the left, click **Apply** to preview it live, and copy the
generated snippet — the fastest way to discover what the editor can do and wire it
into your own app. Built with Vite, like [`embed-live`](../embed-live) and
[`playground`](../playground).

## Run it

From the **repo root**:

```sh
npm install            # once, links the workspace packages
npm run build:lib      # produces frontend/dist-lib the example imports
npm run dev --workspace @cw/example-editor-constructor
```

Then open the URL Vite prints (default `http://localhost:5183/`).

## What to look at — `src/main.ts`

- **One schema drives everything.** A single `SCHEMA` array describes every
  constructor option; the same list builds the form DOM *and* the generated
  snippet, so there's one source of truth. Each field carries explicit `get`/`set`
  closures over the `state` object (no stringly-typed paths), which keeps it
  type-safe even after **Reset** swaps `state` out.
- **The snippet is minimal.** `collect()` diffs the form state against the library
  defaults and emits **only** the options you changed — exactly what you'd paste
  into your app. `buildSnippet()` pretty-prints that as JS object source; the four
  function options (`onShareLink`, `onSave`, `resolveField`, `customizeRibbon`)
  appear as ready-to-fill stubs when toggled on, and `onLoadProgress` is always
  shown (it's how this page drives the loader bar).
- **Apply re-mounts the editor.** Every option is *construction-time only*, so a
  change can only take effect by re-constructing. `rebuild()` calls the previous
  instance's **`destroy()`** (which aborts listeners, closes any collab socket, and
  removes the editor's root) and then `new WordCanvas(...)` into the *same*
  `#editor` element. A generation counter ignores a torn-down instance's late
  `onLoadProgress`, so rapid Apply clicks never flicker.
- **The color "override?" gate.** A `<input type=color>` has no "unset" state and
  the library's true default hexes aren't knowable from the form, so each `theme`
  color is paired with an override checkbox — only checked colors enter the snippet
  and preview, keeping both faithful to the real defaults. The
  **Use darkCanvasTheme preset** button fills the relevant colors from the exported
  `darkCanvasTheme`.

See [`frontend/types/wordcanvas.d.ts`](../../frontend/types/wordcanvas.d.ts) for the
full, authoritative `WordCanvasOptions` reference.
