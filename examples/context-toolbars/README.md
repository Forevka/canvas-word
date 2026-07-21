# Contextual toolbars

[![Open in StackBlitz](https://developer.stackblitz.com/img/open_in_stackblitz.svg)](https://stackblitz.com/github/Forevka/canvas-word/tree/main/examples/context-toolbars)

The editor shows **one floating mini-toolbar for whatever the caret/selection is
on** — built-in bars for a selected **image**, a **hyperlink**, and a **text
selection**, plus any **custom** bars you register with the **`contextToolbars`**
option. A priority-based manager shows only the single most-relevant one. A static,
no-build page (browser import map), like [`embed-offline`](../embed-offline).

## Run it

From the **repo root** (so `/node_modules/...` in the import map resolves):

```sh
npm install            # once, links the workspace packages
npm run build:lib      # produces frontend/dist-lib (wordcanvas.js + builder.js)
npx serve .            # or: python -m http.server
```

Then open `http://localhost:3000/examples/context-toolbars/`. Try:

- **Select a sentence** → the **text format** bar.
- **Caret inside the blue link** (no selection) → the **hyperlink** bar (Open /
  Edit / Copy / Remove).
- **Caret or selection inside the table** → the **custom** bar this page registered.

## Built-in bars

| Bar | Shows when | Priority |
| --- | --- | --- |
| Image | an image is selected | 30 |
| Equation | an equation is selected — edit / align / delete | 29 |
| Table | 2+ table cells are selected (drag across cells) — merge, insert/delete rows & columns | 28 |
| Comment/suggestion | caret inside a comment thread or tracked change — accept/reject · open/resolve | 26 |
| Hyperlink | the caret (collapsed) is inside a link | 25 |
| Footnote/endnote | caret on a note reference marker — go to note / delete | 24 |
| TOC | caret inside a table-of-contents entry — update | 22 |
| Text format | a range is selected (or a caret, if `floatingToolbar.onCaret`) | 20 |
| List item | caret in a bulleted/numbered list — promote / demote / remove | 18 |
| Empty paragraph | caret on an empty line — a ＋ insert menu (heading / list / table / page break / TOC / footnote) | 16 |

The caret-driven bars (comment, footnote/endnote, TOC, list, empty-paragraph — and
hyperlink) show on a *collapsed caret* only, so selecting a text range still gets the
format bar.

Higher priority wins, so a selected image always beats a text selection, and a text
*range* inside a link shows the format bar (not the link bar). All are edit-only.

## The public API — `contextToolbars`

Register your own bars. Each spec has a `when(ctx)` predicate, `buttons`, an optional
`priority`, and an optional `anchor(ctx)` (defaults to the selection):

```js
new WordCanvas({
  container,
  contextToolbars: [
    {
      id: "demo.table",
      priority: 22,                       // above the text bar (20)
      when: (ctx) => ctx.format.inTable,  // ToolbarContext, see below
      buttons: [
        { id: "demo.table.hi", icon: "👋", tooltip: "Insert a wave",
          onClick: (ctx) => ctx.insertText(" 👋 ") },
      ],
    },
  ],
});
```

**`ToolbarContext`** (passed to `when` / `anchor`):

| Field | Meaning |
| --- | --- |
| `format` | `CurrentFormat` at the caret — `imageSelected`, `inTable`, `inContentControl`, bold/italic/…, fontFamily/size, … |
| `selection` | the live `DocSelection`, or `null` |
| `hasRange` | true for a non-empty selection (vs a bare caret) |
| `linkUrl` | the hyperlink URL at the caret, or `null` |
| `selectionRect()` | viewport anchor rect of the selection/caret (the default anchor) |
| `objectRect()` | viewport anchor rect of a selected image, or `null` |

**Buttons** are the same custom-button shape as `floatingToolbar` / the ribbon:
`{ id, icon | label, tooltip?, onClick, active? }`. `onClick(ctx)` receives the editor
handle plus macro helpers (`insertText`, `getDocument`, `setDocument`, `emit`,
`registerCleanup`, …). `active(fmt)` gives the button a pressed state that syncs with
the selection.

See [`app.js`](./app.js) for the full demo.
