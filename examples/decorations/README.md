# Decorations API

[![Open in StackBlitz](https://developer.stackblitz.com/img/open_in_stackblitz.svg)](https://stackblitz.com/github/Forevka/canvas-word/tree/main/examples/decorations)

Draw custom, **paint-only overlays at document coordinates** — highlight a range,
underline or box it, or badge a position — through the **decorations API** on the
editor handle. It never measures text or re-breaks lines: you describe visuals in
document space and the editor resolves them against the live layout (the same feed
the track-changes/comment overlays use). A static, no-build page (browser import
map), like [`embed-offline`](../embed-offline).

## Run it

From the **repo root** (so `/node_modules/...` in the import map resolves):

```sh
npm install
npm run build:lib
npx serve .            # or: python -m http.server
```

Then open `http://localhost:3000/examples/decorations/`.

## What to look at — `app.js`

Decorations live on the `EditorHandle` (get it from `whenReady()`), not the
constructor:

```js
const handle = await editor.whenReady();

handle.setDecorations([
  // fill under the text
  { type: "highlight", range: someSelection, color: "#ffe082", opacity: 0.4 },
  // a line at the baseline
  { type: "underline", range: r, color: "#e53935", thickness: 2 },
  // a stroked rect around the range
  { type: "box", range: r, color: "#1e88e5" },
  // a small marker at a single position
  { type: "badge", at: { blockId, offset }, color: "#8e24aa", label: "1" },
]);

handle.clearDecorations();        // === setDecorations([])
handle.invalidateDecorations();   // re-resolve + repaint (if you mutated a spec in place)
```

- **Document space in, pixels handled for you.** A range decoration takes a
  `DocSelection` (`{ anchor, focus }` — the exact shape `getSelection()` returns);
  a badge takes a `DocPosition` (`{ blockId, offset }`). The editor owns DPR,
  zoom, virtualization, and scroll — you never touch device pixels.
- **Declarative, replace-the-whole-set.** `setDecorations` takes the full array
  each time (like React state); the demo keeps a local `decos` list and re-applies
  it on every change.
- **Anchored across edits.** Decorations are re-resolved after every mutation, so
  a highlight follows its text as you type/delete above it — try it: highlight a
  word, then add a paragraph before it.
- **Interactive (optional).** Give any spec an `onClick` and it becomes clickable —
  hovering shows a pointer cursor and a click fires the handler *instead of*
  placing a caret (like a comment pin). Plain decorations without `onClick` stay
  purely visual and never intercept clicks.

  ```js
  handle.setDecorations([
    { type: "badge", at: pos, color: "#1e88e5", label: "1",
      onClick: (ev) => openAnnotation(ev.clientX, ev.clientY) },
  ]);
  ```

## Try it

The demo draws an initial **interactive** badge + a highlight on load (built
straight from the document model), plus a control panel: select some text and
click **Highlight**, **Underline**, or **Box**; place the caret and click **Badge
at caret** (its badges are clickable too); **Clear all** removes them. Click the
blue “1” badge to see its `onClick` fire in the panel. `window.__handle` is
exposed for console poking.
