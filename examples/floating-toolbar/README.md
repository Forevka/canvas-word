# Floating toolbar

[![Open in StackBlitz](https://developer.stackblitz.com/img/open_in_stackblitz.svg)](https://stackblitz.com/github/Forevka/canvas-word/tree/main/examples/floating-toolbar)

Configure the editor's **floating selection mini-toolbar** (Word's selection
toolbar) through the **`floatingToolbar`** constructor option. A static, no-build
page (browser import map), like [`embed-offline`](../embed-offline).

## Run it

From the **repo root** (so `/node_modules/...` in the import map resolves):

```sh
npm install            # once, links the workspace packages
npm run build:lib      # produces frontend/dist-lib that the import map points at
npx serve .            # or: python -m http.server
```

Then open `http://localhost:3000/examples/floating-toolbar/` (port depends on your
static server). Select some text to see the bar; use the buttons up top to swap
between presets.

## The option — `floatingToolbar`

Pass a boolean to toggle the default set, or an object to customize it:

```js
new WordCanvas({
  container,
  floatingToolbar: {
    enabled: true,   // default true — set false (or pass `false`) to hide it
    onCaret: false,  // default false — true also shows it at a bare caret
    buttons: [       // omit for the full built-in set
      "font", "|", "fontSize", "|",
      "bold", "italic", "underline", "strikethrough", "|",
      "color", "highlight", "|", "clearFormat",
    ],
  },
});
```

**Built-in button ids:** `"font"`, `"fontSize"`, `"bold"`, `"italic"`,
`"underline"`, `"strikethrough"`, `"color"`, `"highlight"`, `"clearFormat"`. Use
`"|"` for a separator. List only the ones you want, in the order you want them.

**Custom buttons.** An entry can also be your own button — the same shape a custom
ribbon button uses, and its `onClick` receives the same `RibbonActionContext`
(editor handle + `insertText` / `getDocument` / `setDocument` / `emit` /
`registerCleanup` …):

```js
buttons: [
  "bold", "italic", "|",
  {
    id: "demo.date",
    icon: "📅",                // SVG string, emoji, or omit and pass `label`
    tooltip: "Insert today's date",
    onClick: (ctx) => ctx.insertText(new Date().toLocaleDateString()),
    active: (fmt) => false,    // optional pressed-state predicate (fmt = CurrentFormat)
  },
]
```

## What to look at — `app.js`

The demo re-mounts the editor with each preset so you can compare them:

- **Default** — `floatingToolbar: true` (the full built-in bar).
- **Minimal** — `{ buttons: ["bold", "italic", "underline"] }`.
- **Reordered, no font picker** — a rearranged subset with separators.
- **Show at caret too** — `{ onCaret: true }` (pops even with no selection).
- **With a custom button** — a 📅 button that inserts today's date via `ctx`.
- **Disabled** — `false` (the ribbon still has every command).

The bar is always edit-only (never shown in view mode) and mutually exclusive with
the floating **image** toolbar.
