# Customizing the ribbon

The `@forevka/wordcanvas` ribbon toolbar is customizable per editor instance via
the `customizeRibbon` constructor option. You can reorder or remove the built-in
tabs/groups/buttons, and add your own — for **macros** (automate document edits),
**config popups**, and **informational popups**.

```ts
import { WordCanvas } from "@forevka/wordcanvas";

new WordCanvas({
  container,
  customizeRibbon(api) {
    // …mutate the ribbon here (see below)…
  },
});
```

`customizeRibbon(api)` runs once at mount, before the ribbon is shown. It is a
no-op when the ribbon is hidden (`view.toolbar: false` or `readonly`).

## The `RibbonApi`

Everything is addressed by **id**. An unknown id is ignored with a console warning
(so your config survives editor upgrades). Discover the current ids — built-in and
custom — at runtime:

```ts
api.tabs();            // ["file","home","insert","layout","table","view", …]
api.groups("home");    // ["home.clipboard","home.font","home.paragraph","home.styles","home.editing"]
api.items("home.font"); // ["home.font.bold","home.font.italic", …]
```

### Reorder / remove built-ins

```ts
api.moveItem("home.font.bold", { after: "home.font.italic" }); // swap Bold/Italic
api.removeItem("home.clipboard.cut");                          // drop a button
api.moveGroup("home.editing", { before: "home.font" });        // reorder a group
api.removeTab("table");                                         // drop a whole tab
api.moveTab("view", { after: "home" });                        // reorder tabs
```

Anchors are `{ before: id }` or `{ after: id }`; omit them to append at the end.

### Built-in id scheme

Ids are namespaced `tab`, `tab.group`, `tab.group.item`, where the `group`/`item`
segments are slugged from each control's label/tooltip (shortcuts in parentheses
are dropped), de-duped within a group with a `-2`, `-3`, … suffix. The tab ids are
the stable set above. Because item ids derive from tooltips, **discover them with
`api.items(groupId)`** rather than hard-coding — that list is the source of truth.

## Adding your own tabs, groups, and buttons

```ts
customizeRibbon(api) {
  // A new tab after View, with one group.
  api.addTab({ id: "myco", label: "My Co.", after: "view" });
  api.addGroup("myco", { id: "myco.macros", label: "Macros" });

  // A macro button: read the doc, transform it, write it back.
  api.addButton("myco.macros", {
    id: "myco.macros.upper",
    label: "UPPER",
    tooltip: "Uppercase the whole document",
    onClick: (ctx) => {
      const doc = ctx.getDocument();
      // …transform a deep clone of `doc`…
      ctx.setDocument(transformed);
    },
  });

  // Add a button into a BUILT-IN tab/group too.
  api.addButton("home.editing", {
    id: "myco.editing.stamp",
    icon: "📌",                       // SVG string, emoji, or text
    tooltip: "Insert today's date",
    onClick: (ctx) => ctx.insertText(new Date().toLocaleDateString()),
  });
}
```

### The button context (`RibbonActionContext`)

`onClick` receives the editor handle plus macro helpers:

- Everything on the `WordCanvas` handle: `getDocument()`, `setDocument()`,
  `openDocx()`, `exportDocx()` / `exportPdf()`, `getMode()` / `setMode()`,
  `getReview()`, comment methods, `share()`, …
- `getSelection()` — the current caret/selection (or `null`).
- `insertText(text)` — insert plain text at the caret.
- `emit(name, payload?)` — fire a `custom` event the embedder can listen for:
  `wc.on("custom", ({ name, payload }) => …)`.
- `registerCleanup(node | fn)` — tie a popup element (removed) or a callback (run)
  to the editor's `destroy()`.

### Optional toggle / enabled state

```ts
api.addButton("home.font", {
  id: "myco.font.highlight",
  icon: "🖍️",
  tooltip: "My highlight",
  active: (fmt) => Boolean(fmt.highlightColor),   // pressed-state, synced on selection change
  enabled: (fmt) => !fmt.imageSelected,           // greyed out when false
  onClick: (ctx) => { /* … */ },
});
```

## Config / informational popups

Bring your own DOM and tie it to teardown, or use the exported `makeFloatingDialog`
helper (a draggable, non-blocking panel):

```ts
import { makeFloatingDialog } from "@forevka/wordcanvas";

api.addButton("myco.macros", {
  id: "myco.macros.settings",
  label: "Settings",
  onClick: (ctx) => {
    const backdrop = document.createElement("div");
    const modal = document.createElement("div");
    const header = document.createElement("div");
    header.className = "cw-drag-handle";
    header.textContent = "My settings";
    // …build your form into `modal`, append `header` first…
    backdrop.append(modal);
    document.body.append(backdrop);
    const ac = new AbortController();
    makeFloatingDialog({ backdrop, modal, handle: header, signal: ac.signal });
    ctx.registerCleanup(backdrop);          // removed on editor destroy
    ctx.registerCleanup(() => ac.abort());  // detaches drag listeners
  },
});
```

## Notes

- **Per-instance.** Several editors on one page can have different ribbons; a custom
  button's `onClick` is scoped to its own editor's handle.
- **Survives document replacement.** The handle in `ctx` always targets the live
  document, so macros keep working after `setDocument()` / opening a `.docx`.
- **Exported types:** `RibbonApi`, `RibbonButtonSpec`, `RibbonActionContext`,
  `CustomizeRibbon`, `FloatingDialogOptions`.
