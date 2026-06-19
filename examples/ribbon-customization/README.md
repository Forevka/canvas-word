# Ribbon customization

Customize the editor's ribbon toolbar through the **`customizeRibbon`** constructor
option: reorder/remove the built-in tabs, groups, and buttons, and add your own —
for **macros**, **config popups**, and **informational popups**. A static,
no-build page (browser import map), like [`embed-offline`](../embed-offline).

## Run it

From the **repo root** (so `/node_modules/...` in the import map resolves):

```sh
npm install            # once, links the workspace packages
npm run build:lib      # produces frontend/dist-lib that the import map points at
npx serve .            # or: python -m http.server
```

Then open `http://localhost:3000/examples/ribbon-customization/` (port depends on
your static server).

## What to look at — `app.js`

The `customizeRibbon(api)` callback runs once at mount with a `RibbonApi`:

- **Discover ids.** Everything is addressed by a stable, namespaced id
  (`tab` · `tab.group` · `tab.group.item`). Rather than hard-code them, inspect at
  runtime — `api.tabs()`, `api.groups("home")`, `api.items("home.font")` (the demo
  logs these to the console).
- **Reorder / remove built-ins.** `api.moveItem("home.font.italic", { before: "home.font.bold" })`
  and `api.removeItem("home.clipboard.cut")`. Unknown ids are ignored with a console
  warning, so your config survives editor upgrades.
- **Add your own.** `api.addTab` / `api.addGroup` / `api.addButton` — the demo adds an
  **Automate** tab (after View) with **Macros** and **Tools** groups, and also adds a
  button into the built-in **Home ▸ Editing** group.

### The button context (macros + popups)

A custom button's `onClick(ctx)` receives the editor handle plus macro helpers. The
demo shows each:

- **Insert at the caret** — `ctx.insertText(...)` (the 📅 date and ✓ stamp buttons).
- **Whole-document macro** — `ctx.getDocument()` → transform a copy → `ctx.setDocument(...)`
  (the **UPPER** button).
- **Informational popup** — read stats with `ctx.getDocument()` and show them in a
  draggable panel built with the exported **`makeFloatingDialog`**, tied to teardown
  via **`ctx.registerCleanup(...)`** (the 🔢 word-count button).
- **Drive the handle API** — `ctx.exportDocx()` (the ⬇️ export button).
- **Emit events** — `ctx.emit(name, payload)` surfaces as the `custom` event
  (`editor.on("custom", …)`); the demo logs every emit to the panel in the corner.

Custom-button `icon` accepts an SVG string, an emoji, or text; pass `label` for a
text button. Optional `active(fmt)` / `enabled(fmt)` predicates sync live with the
selection (the date button greys out while an image is selected).

See [`RIBBON.md`](../../RIBBON.md) for the full API reference.
