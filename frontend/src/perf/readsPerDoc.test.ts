// Step-0 measurement: how many times is a given Document identity read through
// the rebuild-everything data-access layer (paragraphsOf / bandParagraphs) during
// a REAL editing + navigation session? That "reads per doc" number is what decides
// whether the WeakMap index from step 1 is worth shipping — step 1 proved the
// per-call speedup; this proves real workloads hit the winning condition.
//
// Driven through the actual code paths:
//   • edit  = the real `insertText` command + `applyOp` (the DOM-free edit path)
//   • nav   = the selection controller's caret code, copied VERBATIM from
//             src/input/selectionController.ts (it lives in closures and can't be
//             imported); see the citations on each function below.
//
// Inputs: two real appraisal reports + a 5k-paragraph synthetic doc.
//
// Run:  npx vitest run src/perf/readsPerDoc.test.ts   (from frontend/)

import { readdirSync, readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import {
  applyOp,
  BAND_CONTAINERS,
  isHiddenParagraph,
  makeDefaultCharStyle,
  makeDefaultParaStyle,
  navigableParagraphs,
  nextGrapheme,
  paragraphsOf,
  prevGrapheme,
  resetParaScanStats,
  _paraScanStats,
  type Block,
  type DocPosition,
  type DocSelection,
  type Document,
  type Paragraph,
  type SectionProps,
} from "@cw/shared";
import { runImport } from "../import/docx/pipeline";
import { insertText } from "../editor/commands";
import type { EditorState } from "../editor/state";

// ---------------------------------------------------------------------------
// Verbatim copy of the selection controller's body-navigation caret code
// (src/input/selectionController.ts:121-166). Reproduced because those helpers are
// closures over `deps` and aren't exported. Body editing only (getStory() = null),
// logical horizontal step (the common ArrowLeft/Right / Ctrl+arrow path).

const makeNav = (getDoc: () => Document) => {
  // selectionController.ts:121-132 (body branch: memoized navigable-paragraph list)
  const paragraphs = (): Paragraph[] => navigableParagraphs(getDoc());
  // selectionController.ts:138-141
  const textOf = (blockId: string): string => {
    const block = paragraphs().find((b) => b.id === blockId);
    return block ? block.runs.map((r) => r.text).join("") : "";
  };
  // selectionController.ts:143-144
  const blockIndexOf = (blockId: string): number => paragraphs().findIndex((b) => b.id === blockId);
  // selectionController.ts:148-166 (dir = +1, byWord = false: ArrowRight by grapheme)
  const moveRight = (pos: DocPosition): DocPosition => {
    const text = textOf(pos.blockId);
    const blocks = paragraphs();
    const bi = blockIndexOf(pos.blockId);
    if (pos.offset < text.length) {
      return { blockId: pos.blockId, offset: nextGrapheme(text, pos.offset) };
    }
    const next = blocks[bi + 1];
    return next ? { blockId: next.id, offset: 0 } : pos;
  };
  const moveLeft = (pos: DocPosition): DocPosition => {
    const text = textOf(pos.blockId);
    const blocks = paragraphs();
    const bi = blockIndexOf(pos.blockId);
    if (pos.offset > 0) return { blockId: pos.blockId, offset: prevGrapheme(text, pos.offset) };
    const prev = blocks[bi - 1];
    return prev ? { blockId: prev.id, offset: textOf(prev.id).length } : pos;
  };
  return { moveRight, moveLeft };
};

// ---------------------------------------------------------------------------
// 5k-paragraph synthetic doc (same shape as the step-1 bench: ~20% in table cells
// + a few band paragraphs).

const CHAR = makeDefaultCharStyle();
const PARA = makeDefaultParaStyle();
const para = (id: string): Paragraph => ({
  kind: "paragraph",
  id,
  revision: 0,
  runs: [{ text: `text ${id} lorem ipsum dolor sit amet`, style: { ...CHAR } }],
  style: { ...PARA },
});
function makeSyntheticDoc(n: number): Document {
  const blocks: Block[] = [];
  let made = 0;
  let group = 0;
  while (made < n) {
    if (group % 5 === 4) {
      const cellParas: Paragraph[] = [];
      for (let k = 0; k < 3 && made < n; k++, made++) cellParas.push(para(`p${made}`));
      blocks.push({ kind: "table", id: `t${group}`, revision: 0, rows: [{ cells: [{ id: `c${group}`, blocks: cellParas }] }] });
    } else {
      blocks.push(para(`p${made}`));
      made++;
    }
    group++;
  }
  const section: SectionProps = {
    pageWidthPx: 816,
    pageHeightPx: 1056,
    marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
    header: [para("hdr-0"), para("hdr-1")],
    footer: [para("ftr-0"), para("ftr-1")],
  };
  return { section, blocks };
}

// Real reports are developer-local artifacts: `.docx` files dropped into the
// frontend/ root, which is gitignored for *.docx (they carry real addresses/PII,
// so neither the binaries NOR their identifying filenames belong in the repo). We
// DISCOVER whatever local `.docx` are present at runtime (no PII in source) and
// measure them; in CI there are none, so only the synthetic case runs. Word lock
// files (`~$…`) are skipped.
const frontendRoot = fileURLToPath(new URL("../../", import.meta.url));
const localReportFiles = (): string[] => {
  try {
    return readdirSync(frontendRoot).filter((f) => f.toLowerCase().endsWith(".docx") && !f.startsWith("~$"));
  } catch {
    return [];
  }
};
const docOfFile = (file: string): Document =>
  runImport(new Uint8Array(readFileSync(`${frontendRoot}${file}`))).doc;

// ---------------------------------------------------------------------------
// Workloads. Each returns the tallied counters + wall-clock.

const firstBodyParaId = (doc: Document): string => {
  for (const p of paragraphsOf(doc)) {
    if (!isHiddenParagraph(p)) return p.id;
  }
  throw new Error("no body paragraph");
};

interface Tally {
  fullCalls: number;
  fullSlots: number;
  bandCalls: number;
  bandSlots: number;
  ms: number;
  steps: number;
}
function measure(steps: number, run: () => void): Tally {
  resetParaScanStats();
  _paraScanStats.enabled = true;
  const t0 = performance.now();
  run();
  const ms = performance.now() - t0;
  _paraScanStats.enabled = false;
  return {
    fullCalls: _paraScanStats.fullCalls,
    fullSlots: _paraScanStats.fullSlots,
    bandCalls: _paraScanStats.bandCalls,
    bandSlots: _paraScanStats.bandSlots,
    ms,
    steps,
  };
}

// EDIT: type `keystrokes` characters one at a time via the real insertText command.
// Each keystroke produces a NEW doc identity (the cache-miss path).
function editSession(doc: Document, keystrokes: number): { docs: number } & Tally {
  let state: EditorState = { doc, selection: collapsed(firstBodyParaId(doc), 0) };
  const seen = new Set<Document>();
  const t = measure(keystrokes, () => {
    for (let i = 0; i < keystrokes; i++) {
      const trn = insertText("x")(state);
      if (!trn) continue;
      let d = state.doc;
      for (const op of trn.ops) d = applyOp(d, op).doc;
      seen.add(d);
      state = { doc: d, selection: trn.selectionAfter };
    }
  });
  return { ...t, docs: seen.size };
}

// NAV: step the caret right `steps` times over a STABLE doc (no edits) — arrow-key
// navigation / repeated reads of the same identity (the cache-hit path).
function navSession(doc: Document, steps: number): Tally {
  const { moveRight } = makeNav(() => doc);
  let pos: DocPosition = { blockId: firstBodyParaId(doc), offset: 0 };
  return measure(steps, () => {
    for (let i = 0; i < steps; i++) pos = moveRight(pos);
  });
}

const collapsed = (blockId: string, offset: number): DocSelection => ({
  anchor: { blockId, offset },
  focus: { blockId, offset },
});

// ---------------------------------------------------------------------------

const fmt = (n: number): string => n.toLocaleString("en-US");
const pad = (s: string, w: number): string => s.padEnd(w);
const padN = (n: string, w: number): string => n.padStart(w);

function report(label: string, doc: Document): { nav: Tally; steps: number } {
  const size = navigableParagraphs(doc).length;
  const KEYS = 40;
  const STEPS = 100;
  const edit = editSession(doc, KEYS);
  const nav = navSession(doc, STEPS);

  const lines: string[] = [];
  lines.push("");
  lines.push(`### ${label}  —  ${fmt(size)} navigable paragraphs`);
  lines.push(
    pad("workload", 22) +
      padN("fullScans", 11) +
      padN("/event", 9) +
      padN("slotsWalked", 14) +
      padN("/event", 10) +
      padN("bandScans", 12) +
      padN("ms", 9),
  );
  const row = (name: string, t: Tally) =>
    pad(name, 22) +
    padN(fmt(t.fullCalls), 11) +
    padN((t.fullCalls / t.steps).toFixed(1), 9) +
    padN(fmt(t.fullSlots), 14) +
    padN(fmt(Math.round(t.fullSlots / t.steps)), 10) +
    padN(fmt(t.bandCalls), 12) +
    padN(t.ms.toFixed(1), 9);
  lines.push(row(`edit (${KEYS} keys)`, edit));
  lines.push(row(`nav (${STEPS} arrows)`, nav));
  lines.push(
    `   edit produced ${edit.docs} distinct doc identities; ` +
      `nav kept 1 stable identity across ${STEPS} steps.`,
  );
  // eslint-disable-next-line no-console
  console.log(lines.join("\n"));
  return { nav, steps: STEPS };
}

describe("step-0: reads per document identity", () => {
  // Measure any developer-local `.docx` reports that happen to be present, with
  // anonymized labels (filenames come from the filesystem at runtime, never from
  // committed source). Empty in CI — only the synthetic guard below runs there.
  const locals = localReportFiles();
  locals.forEach((file, i) => {
    it(`local report #${i + 1}`, () => {
      report(`local report #${i + 1}`, docOfFile(file));
    });
  });

  // Committed regression guard: with the per-doc memoization in place, caret
  // navigation over a large stable doc must stay O(1) — after the first read
  // warms the index, 100 arrow keys do ZERO full-document rebuilds, walk ZERO
  // paragraph slots, and build the band-id set at most once. Reverting the
  // memoization would balloon these into the millions (see PR description).
  it("synthetic 5k — caret navigation stays O(1) per keystroke", () => {
    const { nav, steps } = report("synthetic 5k", makeSyntheticDoc(5000));
    expect(steps).toBe(100);
    expect(nav.fullCalls).toBe(0); // paragraphsOf never rebuilds during navigation
    expect(nav.fullSlots).toBe(0); // …so no paragraph slots are walked
    expect(nav.bandCalls).toBeLessThanOrEqual(BAND_CONTAINERS.length); // band set built once
  });
});
