// Contextual floating toolbars demo. The editor shows ONE floating bar for whatever
// the caret/selection is on — built-in bars for a selected image, a hyperlink, and a
// text selection, PLUS any you register via the `contextToolbars` option. This demo
// seeds a document (hyperlink, table, list, equation, footnote) and registers a custom
// TABLE toolbar.
import { WordCanvas } from "@forevka/wordcanvas";
import { DocumentBuilder } from "@forevka/wordcanvas/builder";

// --- seed a document with a hyperlink and a table ------------------------------
const doc = DocumentBuilder.create()
  .paragraph("Contextual toolbars", { bold: true, fontSizePx: 28 })
  .paragraph("Select this whole sentence to see the text format bar appear above it.")
  .paragraph("Put your caret inside ")
  .text("this hyperlink", { link: "https://forevka.dev" })
  .text(" (no selection) to see the hyperlink bar — Open / Edit / Copy / Remove.")
  .paragraph("Click a cell in the table below to see a CUSTOM toolbar that this page registered with the contextToolbars option:")
  .table([["Item", "Qty"], ["Widget", "3"], ["Gadget", "7"]], { headerRow: true })
  .paragraph("Click into this list to see the list toolbar (promote / demote / remove):")
  .bulletList(["First item", "Second item", "Third item"])
  .paragraph("Select this equation to see the equation toolbar (edit / align / delete):")
  .equation("E = mc^2")
  .paragraph("This sentence has a footnote")
  .footnote("Put the caret on the footnote marker to see the note bar — Go to note / Delete.")
  .text(" — click its superscript marker to see the footnote bar.")
  .paragraph("Put the caret on the empty line below to see the ＋ insert menu:")
  .paragraph("")
  .build();

// --- first-load progress bar ---------------------------------------------------
const loader = document.getElementById("loader");
const progress = loader.querySelector(".bar");
const label = loader.querySelector(".label");
const PHASE_LABEL = { bundle: "Loading editor…", fonts: "Loading fonts…", ready: "Ready" };

// --- mount with a custom context toolbar --------------------------------------
const editor = new WordCanvas({
  container: document.getElementById("editor"),
  onLoadProgress: ({ phase, percent }) => {
    progress.style.width = `${Math.round(percent * 100)}%`;
    label.textContent = PHASE_LABEL[phase] ?? "";
    if (phase === "ready") loader.classList.add("done");
  },

  // Register a custom contextual toolbar. It shows (above the caret/selection)
  // whenever `when(ctx)` is true and no higher-priority bar is active. Its buttons
  // are the same custom-button shape the ribbon and floatingToolbar use — `onClick`
  // gets the editor handle + macro helpers (insertText / getDocument / emit / …).
  contextToolbars: [
    {
      id: "demo.table",
      priority: 22, // above the text format bar (20), so it wins inside a table
      when: (ctx) => ctx.format.inTable, // ToolbarContext: format/selection/hasRange/linkUrl
      buttons: [
        {
          id: "demo.table.hi",
          icon: "👋",
          tooltip: "Insert a wave into this cell",
          onClick: (ctx) => ctx.insertText(" 👋 "),
        },
        {
          id: "demo.table.count",
          icon: "🔢",
          tooltip: "Count the blocks in the document",
          onClick: (ctx) => {
            const count = (blocks) =>
              blocks.reduce(
                (n, b) =>
                  n + 1 + (b.kind === "table" ? b.rows.flatMap((r) => r.cells).reduce((m, c) => m + count(c.blocks), 0) : 0),
                0,
              );
            window.alert(`This document has ${count(ctx.getDocument().blocks)} blocks.`);
          },
        },
      ],
    },
  ],
});
window.__wc = editor;

editor.whenReady().then(() => editor.setDocument(doc));
