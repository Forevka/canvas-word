// Merge documents demo. Builds a few independently-authored PARTS with the
// DocumentBuilder (given overlapping ids/bookmarks on purpose), folds them into one
// report with the headless merge API (@forevka/wordcanvas/query — DocumentEditor.append
// / mergeAll), shows the result in a live editor on the right, and applies a
// per-section footer (a borderless [logo | address] table + a "Page X of Y" field)
// with setSectionFooter. Every button runs a real API call and re-renders.

import { WordCanvas } from "@forevka/wordcanvas";
import { DocumentBuilder } from "@forevka/wordcanvas/builder";
import {
  DocumentEditor,
  getSections,
  mergeAll,
  type Block,
  type Document,
  type MergeResult,
} from "@forevka/wordcanvas/query";

const $ = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T;

// ---------------------------------------------------------------------------
// 1. The parts. Same idSeed → identical id spaces (maximal collisions); every part
//    defines a `intro` bookmark, so merging renames the clashes. Distinct page
//    sizes keep each part's section boundary observable after a round-trip.

const part = (title: string, body: string, pageSize: "Letter" | "A4" | "Legal") =>
  DocumentBuilder.create({ idSeed: "part", pageSize })
    .style({ id: "Heading1", name: "Heading 1", char: { bold: true, fontSizePx: 24, color: "#1a1a2e" }, para: { spaceBeforePx: 16, spaceAfterPx: 6 } })
    .paragraph(title).withStyle("Heading1")
    .paragraph(body)
    .paragraph("See ").bookmark("intro", "the introduction").text(" for context.")
    .build();

const cover = (): Document => part("Cover", "The assembled report — built from separate parts.", "Letter");
const body = (): Document => part("Body", "Beta section body. Lorem ipsum dolor sit amet.", "A4");
const appendix = (): Document => part("Appendix", "Supporting tables and notes.", "Legal");

// ---------------------------------------------------------------------------
// 2. A branded content footer as a standalone Block[] (a throwaway builder's
//    footer band): [ logo | right-aligned address ] over a centered page number.

function logoDataUrl(): string {
  const c = document.createElement("canvas");
  c.width = 96;
  c.height = 32;
  const ctx = c.getContext("2d")!;
  ctx.fillStyle = "#1a73e8";
  ctx.fillRect(0, 0, 96, 32);
  ctx.fillStyle = "#fff";
  ctx.font = "bold 15px system-ui, sans-serif";
  ctx.fillText("ACME", 10, 21);
  return c.toDataURL("image/png");
}

function contentFooter(): Block[] {
  const FONT = "Times New Roman";
  const SIZE = 16; // px (~12pt)
  const flush = { before: 0, after: 0 } as const;
  const cell = { borders: {}, margin: { top: 0, right: 2, bottom: 0, left: 2 }, vAlign: "center" as const };
  const logo = logoDataUrl();

  // The TS builder's `.paragraph()` returns a chainable ParagraphBuilder (there is
  // no callback form — unlike the C# facade); `.text().pageField()` etc. chain on it.
  return DocumentBuilder.create()
    .footer((f) => {
      f.paragraph().spacing(flush); // leading spacer
      f.table(
        (t) => {
          t.colFractions([0.28, 0.72]);
          t.row((r) => {
            r.cell((c) => { c.image(logo, { widthPx: 96, heightPx: 32, align: "left" }); }, { ...cell, preferredWidth: { px: 28, type: "pct" } });
            r.cell(
              (c) => {
                c.paragraph("APR-2026-00042").align("right").font(FONT).fontSize(SIZE).spacing(flush);
                c.paragraph("123 Main St, Springfield, IL 62704").align("right").font(FONT).fontSize(SIZE).spacing(flush);
              },
              { ...cell, preferredWidth: { px: 72, type: "pct" } },
            );
          });
        },
        { widthMode: "autofitWindow" }, // cells set their own (empty) borders ⇒ no lines
      );
      f.paragraph().align("center").font(FONT).fontSize(SIZE).spacing({ before: 7, after: 0 }).text("Page ").pageField().text(" of ").numPagesField();
    })
    .build().section.footer ?? [];
}

// ---------------------------------------------------------------------------
// 3. Editor (live preview) + a persistent DocumentEditor over the model.

const wc = new WordCanvas({ container: $("editor") });
let editor = new DocumentEditor(cover());
let lastMerge: MergeResult | null = null;

const note = (msg: string | null): void => {
  const el = $("note");
  el.hidden = msg === null;
  el.textContent = msg ?? "";
};

async function render(): Promise<void> {
  const doc = editor.doc;
  const lines: string[] = [];
  const add = (label: string, value: unknown): void => {
    lines.push(`${label.padEnd(20)} ${value}`);
  };

  lines.push("── DOCUMENT ──────────────────────────────────");
  add("blocks", doc.blocks.length);
  add("sections", getSections(doc).length);
  add("bookmarks", Object.keys(doc.bookmarks ?? {}).join(", ") || "—");
  add("can undo / redo", `${editor.canUndo} / ${editor.canRedo}`);

  if (lastMerge) {
    lines.push("");
    lines.push("── LAST MERGE ────────────────────────────────");
    add("blocks remapped", Object.keys(lastMerge.idMap.blocks).length);
    add("bookmarks remapped", Object.keys(lastMerge.idMap.bookmarks).length || "—");
    add("warnings", lastMerge.warnings.map((w) => `${w.code}${w.detail ? `(${w.detail})` : ""}`).join(", ") || "none");
  }

  $("report").textContent = lines.join("\n");
  await wc.setDocument(doc);
}

// ---------------------------------------------------------------------------
// 4. Actions — each is a real merge/edit API call, then re-render.

interface Action {
  label: string;
  primary?: boolean;
  run: () => void;
}

const actions: Action[] = [
  {
    label: "Append Body (new page)",
    primary: true,
    run: () => {
      lastMerge = editor.append(body(), { sectionBreak: "nextPage" });
    },
  },
  {
    label: "Append Appendix",
    run: () => {
      lastMerge = editor.append(appendix(), { sectionBreak: "nextPage" });
    },
  },
  {
    label: "Merge all at once (mergeAll)",
    run: () => {
      const merged = mergeAll([cover(), body(), appendix()], { sectionBreak: "nextPage" });
      lastMerge = merged;
      editor = new DocumentEditor(merged.doc);
    },
  },
  {
    label: "Add footer to last section",
    run: () => editor.setSectionFooter(getSections(editor.doc).length - 1, contentFooter()),
  },
  { label: "Undo", run: () => editor.undo() },
  { label: "Redo", run: () => editor.redo() },
  {
    label: "Reset",
    run: () => {
      editor = new DocumentEditor(cover());
      lastMerge = null;
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
    render().catch((e) => note(`render: ${(e as Error).message}`));
  });
  bar.appendChild(btn);
}

render().catch((e) => note(`render: ${(e as Error).message}`));
