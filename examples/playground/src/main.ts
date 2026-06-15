// Document-builder playground: write builder code on the left, feed it a JSON
// data model, and watch the document render live in the editor on the right.
// The code pane is the body of `async (data, ctx) => Document`; on every code
// or data change (debounced) it re-runs and the result swaps into the editor
// via WordCanvas.setDocument — the declarative-rebuild data-binding model.
//
// Executing the user's code with AsyncFunction is deliberate: this is a dev
// tool running the developer's own code in their own page, not a security
// boundary.

import { WordCanvas } from "@forevka/wordcanvas";
import {
  DocumentBuilder,
  PAGE_SIZES,
  bytesToDataUrl,
  cm,
  inches,
  pt,
  twips,
  type BuilderWarning,
  type CreateOptions,
  type Document,
  type TemplateOptions,
} from "@forevka/wordcanvas/builder";
import { EditorView, basicSetup } from "codemirror";
import { javascript } from "@codemirror/lang-javascript";

const $ = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T;

// ---------------------------------------------------------------------------
// Default content (used when localStorage is empty)

const DEFAULT_CODE = `// Runs as: async (data, ctx) => Document
// ctx = { DocumentBuilder, template, inches, cm, pt, twips, PAGE_SIZES, bytesToDataUrl }
const { DocumentBuilder, template } = ctx;

const b = template
  ? await DocumentBuilder.fromTemplate(template) // styles/page setup/bands from the .docx
  : DocumentBuilder.create();

b.paragraph(data.title).withStyle("Heading1")
 .paragraph(\`Prepared for \${data.customer.name} on \${data.date}\`).italic()
 .paragraph("Line items").withStyle("Heading2")
 .table([
   ["Item", "Qty", "Price"],
   ...data.items.map(i => [i.name, { text: String(i.qty), align: "right" }, { text: "$" + i.price.toFixed(2), align: "right" }]),
 ], { headerRow: true, colFractions: [3, 1, 1] })
 .paragraph("Notes").withStyle("Heading2")
 .bulletList(data.notes)
 .footer(f => f.paragraph("Page {page} of {pages}").align("center").fontSize(11));

return b.build();
`;

const DEFAULT_DATA = `{
  "title": "Invoice #1042",
  "date": "2026-06-11",
  "customer": { "name": "Acme Corp" },
  "items": [
    { "name": "Widget", "qty": 3, "price": 19.5 },
    { "name": "Gadget", "qty": 1, "price": 99 },
    { "name": "Gizmo subscription (annual)", "qty": 2, "price": 240 }
  ],
  "notes": ["Net 30 payment terms.", "Thanks for your business!"]
}`;

// ---------------------------------------------------------------------------
// Panes

const LS_CODE = "cw-playground-code";
const LS_DATA = "cw-playground-data";

const dataPane = $<HTMLTextAreaElement>("data");
dataPane.value = localStorage.getItem(LS_DATA) ?? DEFAULT_DATA;

const codeView = new EditorView({
  parent: $("code"),
  doc: localStorage.getItem(LS_CODE) ?? DEFAULT_CODE,
  extensions: [
    basicSetup,
    javascript(),
    EditorView.updateListener.of((u) => {
      if (u.docChanged) scheduleRun();
    }),
  ],
});

dataPane.addEventListener("input", scheduleRun);

// ---------------------------------------------------------------------------
// Template upload

let template: ArrayBuffer | null = null;

const templateName = $("template-name");
const templateClear = $<HTMLButtonElement>("template-clear");
const templateInput = $<HTMLInputElement>("template-input");
$("template-btn").addEventListener("click", () => templateInput.click());
templateInput.addEventListener("change", async () => {
  const file = templateInput.files?.[0];
  if (!file) return;
  template = await file.arrayBuffer();
  templateName.textContent = file.name;
  templateClear.hidden = false;
  scheduleRun();
});
templateClear.addEventListener("click", () => {
  template = null;
  templateInput.value = "";
  templateName.textContent = "no template";
  templateClear.hidden = true;
  scheduleRun();
});

// ---------------------------------------------------------------------------
// Status strip

const status = $("status");
const showStatus = (kind: "ok" | "warn" | "error", text: string): void => {
  status.className = kind;
  status.textContent = text;
};

// ---------------------------------------------------------------------------
// Editor (offline — the preview needs no backend)

// `agentTools` exposes the previewed document to AI agents over WebMCP — handy for
// inspecting the built document's layout (inspect_layout) while iterating. Note the
// rebuild loop owns the document, so an agent's direct edits are replaced on the
// next code/data change; read & inspect tools are the natural fit here.
const wc = new WordCanvas({ container: $("editor"), agentTools: true });

// ---------------------------------------------------------------------------
// Rebuild loop

// eslint-disable-next-line @typescript-eslint/no-unsafe-function-type
const AsyncFunction = Object.getPrototypeOf(async () => {}).constructor as new (
  ...args: string[]
) => (data: unknown, ctx: unknown) => Promise<unknown>;

let timer: ReturnType<typeof setTimeout> | undefined;
let running = false;
let queued = false;

function scheduleRun(): void {
  clearTimeout(timer);
  timer = setTimeout(() => void run(), 400);
}

async function run(): Promise<void> {
  if (running) {
    queued = true; // coalesce: re-run once the current pass finishes
    return;
  }
  running = true;
  try {
    const code = codeView.state.doc.toString();
    const dataText = dataPane.value;
    localStorage.setItem(LS_CODE, code);
    localStorage.setItem(LS_DATA, dataText);

    let data: unknown;
    try {
      data = JSON.parse(dataText);
    } catch (e) {
      showStatus("error", `Data is not valid JSON: ${e instanceof Error ? e.message : String(e)}`);
      return; // keep the last good preview
    }

    // Track builders the user's code creates so their warnings surface here.
    const builders: { warnings: readonly BuilderWarning[] }[] = [];
    const track = <T extends { warnings: readonly BuilderWarning[] }>(b: T): T => {
      builders.push(b);
      return b;
    };
    const ctx = {
      DocumentBuilder: {
        create: (opts?: CreateOptions) => track(DocumentBuilder.create(opts)),
        fromTemplate: async (docx: ArrayBuffer | Uint8Array, opts?: TemplateOptions) =>
          track(await DocumentBuilder.fromTemplate(docx, opts)),
      },
      template: template ? template.slice(0) : null,
      inches,
      cm,
      pt,
      twips,
      PAGE_SIZES,
      bytesToDataUrl,
    };

    const t0 = performance.now();
    let doc: unknown;
    try {
      doc = await new AsyncFunction("data", "ctx", code)(data, ctx);
    } catch (e) {
      showStatus("error", e instanceof Error ? `${e.name}: ${e.message}` : String(e));
      return;
    }
    if (!doc || typeof doc !== "object" || !Array.isArray((doc as Document).blocks)) {
      showStatus("error", "The code must return the built Document — end with `return b.build()`.");
      return;
    }

    await wc.setDocument(doc as Document);
    const ms = (performance.now() - t0).toFixed(0);
    const warnings = builders.flatMap((b) => b.warnings);
    if (warnings.length > 0) {
      showStatus("warn", `Rebuilt in ${ms}ms with ${warnings.length} warning(s):\n` + warnings.map((w) => `• ${w.message}`).join("\n"));
    } else {
      showStatus("ok", `Rebuilt in ${ms}ms · ${(doc as Document).blocks.length} block(s)`);
    }
  } finally {
    running = false;
    if (queued) {
      queued = false;
      scheduleRun();
    }
  }
}

// First render once the editor is mounted.
void wc.whenReady().then(() => void run());

// Expose for in-browser poking.
(window as unknown as { __playground?: unknown }).__playground = { wc, run };
