// Query & Edit API demo. Builds a document with the DocumentBuilder, shows it in
// a live editor on the right, and drives the HEADLESS query/edit API
// (@forevka/wordcanvas/query) from the panel on the left — the same isomorphic
// core you'd use in Node, a worker, or the C# bindings. Every button runs a real
// API call against the current document and re-renders both the report and the
// editor (via WordCanvas.setDocument), so you can watch edits land.

import { WordCanvas } from "@forevka/wordcanvas";
import { DocumentBuilder, type Document } from "@forevka/wordcanvas/builder";
import {
  DocumentEditor,
  findParagraphs,
  getPages,
  getParagraphs,
  getSdts,
  getSections,
  getStyles,
  positionOfText,
  rangeText,
  sdtText,
} from "@forevka/wordcanvas/query";

const $ = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T;

// ---------------------------------------------------------------------------
// 1. Build a sample document (the thing we'll query & edit).

function sampleDoc(): Document {
  return DocumentBuilder.create({ pageSize: "Letter" })
    .style({ id: "Heading1", name: "Heading 1", char: { bold: true, fontSizePx: 24, color: "#1a1a2e" }, para: { spaceBeforePx: 16, spaceAfterPx: 6 } })
    .style({ id: "Heading2", name: "Heading 2", char: { bold: true, fontSizePx: 18, color: "#1a1a2e" }, para: { spaceBeforePx: 12, spaceAfterPx: 4 } })
    .paragraph("Query & Edit demo", { bold: true, fontSizePx: 30, color: "#1a1a2e" })
    .paragraph("Introduction").withStyle("Heading1")
    .paragraph("Lorem ipsum dolor sit amet. Here is lorem again, and one more lorem to find & replace.")
    .paragraph("Filled by a content control → ")
    .plainTextControl("Jane Doe", { alias: "Full name", tag: "fullName" })
    .text(" ← edit it headlessly with setSdtText.")
    .paragraph("Details").withStyle("Heading2")
    .paragraph("The panel on the left calls the query API on this exact document. Try the edit buttons.")
    .build();
}

// ---------------------------------------------------------------------------
// 2. Editor (live preview) + a persistent DocumentEditor over the model.

const wc = new WordCanvas({ container: $("editor") });
let editor = new DocumentEditor(sampleDoc());

const note = (msg: string | null): void => {
  const el = $("note");
  el.hidden = msg === null;
  el.textContent = msg ?? "";
};

/** Re-run the query API against the current document and paint the report;
 *  push the current document into the editor so edits are visible. */
async function render(): Promise<void> {
  const doc = editor.doc;
  const lines: string[] = [];
  const add = (label: string, value: unknown): void => {
    lines.push(`${label.padEnd(18)} ${value}`);
  };

  lines.push("── QUERY (read-only) ─────────────────────────");
  add("paragraphs", getParagraphs(doc).length);
  add("sections", getSections(doc).length);
  add("styles", getStyles(doc).map((s) => s.name).join(", ") || "—");
  add("find 'lorem'", `${findParagraphs(doc, "lorem").length} paragraph(s)`);

  const sdts = getSdts(doc);
  add("content controls", `${sdts.length}`);
  for (const s of sdts) add(`  · ${s.props.alias ?? s.id}`, `tag=${s.props.tag ?? "—"} text="${sdtText(doc, s.id)}"`);

  const pos = positionOfText(doc, "lorem");
  add("positionOf 'lorem'", pos ? `${pos.blockId} @ ${pos.offset}` : "not found");
  if (pos) add("rangeText (8 ch)", JSON.stringify(rangeText(doc, { anchor: pos, focus: { blockId: pos.blockId, offset: pos.offset + 8 } })));

  // getPages runs a layout pass — needs text measurement, which the mounted
  // editor initializes. Guarded so the rest of the report always renders.
  try {
    add("pages", getPages(doc).length);
  } catch (e) {
    add("pages", `(needs a layout context: ${(e as Error).message})`);
  }

  lines.push("");
  lines.push("── EDIT (undo/redo via the op engine) ────────");
  add("can undo / redo", `${editor.canUndo} / ${editor.canRedo}`);

  $("report").textContent = lines.join("\n");
  await wc.setDocument(doc);
}

// ---------------------------------------------------------------------------
// 3. Actions — each is a real query/edit API call, then re-render.

interface Action {
  label: string;
  primary?: boolean;
  run: () => void;
}

const actions: Action[] = [
  {
    label: "Fill the content control",
    primary: true,
    run: () => {
      const sdt = getSdts(editor.doc)[0];
      if (sdt) editor.setSdtText(sdt.id, "Ada Lovelace");
    },
  },
  {
    label: "Replace all 'lorem' → 'LOREM'",
    run: () => editor.replaceAllText("lorem", "LOREM"),
  },
  {
    label: "Move 'Introduction' down",
    run: () => {
      const hit = findParagraphs(editor.doc, "Introduction")[0];
      if (hit) editor.moveBlock(hit.paragraph.id, 4);
    },
  },
  { label: "Undo", run: () => editor.undo() },
  { label: "Redo", run: () => editor.redo() },
  {
    label: "Reset",
    run: () => {
      editor = new DocumentEditor(sampleDoc());
    },
  },
];

const bar = $("actions");
for (const a of actions) {
  const btn = document.createElement("button");
  btn.textContent = a.label;
  if (a.primary) btn.className = "primary";
  btn.addEventListener("click", () => {
    try {
      a.run();
      note(null);
    } catch (e) {
      note(`${a.label}: ${(e as Error).message}`);
    }
    void render();
  });
  bar.appendChild(btn);
}

void render();
