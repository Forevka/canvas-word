# Command + keymap registry

[![Open in StackBlitz](https://developer.stackblitz.com/img/open_in_stackblitz.svg)](https://stackblitz.com/github/Forevka/canvas-word/tree/main/examples/command-registry)

Register custom **commands** and bind **keyboard shortcuts** to them through the
**`commands`** constructor option — the additive, fork-free way to add editor
behavior. A static, no-build page (browser import map), like
[`embed-offline`](../embed-offline).

## Run it

From the **repo root** (so `/node_modules/...` in the import map resolves):

```sh
npm install            # once, links the workspace packages
npm run build:lib      # produces frontend/dist-lib that the import map points at
npx serve .            # or: python -m http.server
```

Then open `http://localhost:3000/examples/command-registry/` (port depends on your
static server).

## What to look at — `app.js`

The whole feature is one option: an array of commands.

```js
new WordCanvas({
  container,
  commands: [
    {
      id: "demo.insertDate",         // stable, namespaced id
      label: "Insert today's date",  // for docs / a future palette
      keybinding: "Alt+Shift+D",     // one chord, or an array of chords
      run: (ctx) => ctx.insertText(new Date().toLocaleDateString()),
    },
  ],
});
```

- **The handler `ctx`** is the *same* `CommandContext` a custom ribbon button gets
  (the full `EditorHandle` plus `getSelection` / `insertText` / `emit` /
  `registerCleanup`). So a macro is interchangeable between a keystroke, a ribbon
  button, and `handle.runCommand(id)`. The demo reuses the exact
  uppercase-document and word-count macros from the ribbon example.
- **Keybindings** use `Mod+Shift+K` syntax. Modifiers: `Mod`, `Ctrl`, `Alt`,
  `Shift`, `Meta`. **`Mod` = Ctrl on Windows/Linux, ⌘ (Cmd) on macOS**, so one
  binding works on every OS. The key is the final segment, matched against
  `KeyboardEvent.key` (`"k"`, `"Enter"`, `"F2"`, `"/"`, …).
- **Built-ins always win.** The registry never overrides the core editing chords
  (Ctrl+B/I/U/Z/Y, Ctrl+Enter) — it only fires a custom binding the built-in
  keymap didn't already handle.
- **Conflicts are reported, not silent.** A duplicate command `id` or two commands
  claiming the same chord is dropped with a `console.warn` (first registration
  wins), so behavior stays deterministic.

## Invoke programmatically

Every registered command is also callable through the handle:

```js
const handle = await editor.whenReady();
handle.runCommand("demo.insertDate"); // → true, or false if no such id
```

The demo wires a button to `handle.runCommand("demo.insertDate")` and exposes
`window.__handle` so you can try `__handle.runCommand("demo.wordCount")` in the
console.

## Try it

- <kbd>Alt+Shift+U</kbd> — uppercase the whole document
- <kbd>Alt+Shift+D</kbd> — insert today's date at the caret
- <kbd>Alt+Shift+M</kbd> (or <kbd>Alt+Shift+R</kbd>) — insert a `[reviewed]` marker
- <kbd>Ctrl/Cmd+Shift+L</kbd> — word count (emitted to the event log)

Each command calls `ctx.emit(...)`, which surfaces as the standard `custom`
WordCanvas event; the demo logs every one to the panel in the corner.
