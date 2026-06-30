// Step-1 measurement + regression guard for the document data-access layer.
//
// `blockById`/`blockIndexOf` from @cw/shared are NOW memoized (a
// WeakMap<Document, …> id index — see shared/src/model/text.ts), so they can't
// serve as their own "before". This file therefore inlines the PRE-memoization
// behaviour (`*Baseline` below: rebuild the full paragraph traversal via the
// un-memoized `bodyParagraphs`/`bandParagraphs`, then linear-scan) and benches it
// against the SHIPPED memoized helpers. That shows the win the memoization bought
// AND flags any regression that reintroduces per-call rebuilds (the shipped side
// would drift back toward the baseline).
//
// Run:  npx vitest bench src/perf/dataAccess.bench.ts   (from frontend/)

import { bench, describe } from "vitest";
import {
  BAND_CONTAINERS,
  bandParagraphs,
  blockById,
  blockIndexOf,
  bodyParagraphs,
  makeDefaultCharStyle,
  makeDefaultParaStyle,
  type Block,
  type Document,
  type Paragraph,
  type SectionProps,
} from "@cw/shared";

// ---------------------------------------------------------------------------
// Fixture: a doc with N body paragraphs, ~20% of them nested one level deep in
// table cells (so the traversal recurses, like a real document), plus a few
// header/footer band paragraphs (bands are walked on every paragraphsOf call).

const CHAR = makeDefaultCharStyle();
const PARA = makeDefaultParaStyle();

const para = (id: string): Paragraph => ({
  kind: "paragraph",
  id,
  revision: 0,
  runs: [{ text: `text ${id} lorem ipsum dolor sit amet`, style: { ...CHAR } }],
  style: { ...PARA },
});

const SECTION: SectionProps = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
  // a couple of band paragraphs so paragraphsOf's band walk isn't a no-op
  header: [para("hdr-0"), para("hdr-1")],
  footer: [para("ftr-0"), para("ftr-1")],
};

/** N paragraph-bearing blocks: every 5th group of 3 paragraphs is packed into a
 *  1x1 table cell (≈20% of paragraphs live in cells). Returns the doc + the full
 *  ordered list of paragraph ids that paragraphsOf would surface (body only). */
function makeDoc(n: number): { doc: Document; ids: string[] } {
  const blocks: Block[] = [];
  const ids: string[] = [];
  let made = 0;
  let group = 0;
  while (made < n) {
    if (group % 5 === 4) {
      // a table holding up to 3 paragraphs in one cell
      const cellParas: Paragraph[] = [];
      for (let k = 0; k < 3 && made < n; k++, made++) {
        const p = para(`p${made}`);
        cellParas.push(p);
        ids.push(p.id);
      }
      blocks.push({
        kind: "table",
        id: `t${group}`,
        revision: 0,
        rows: [{ cells: [{ id: `c${group}`, blocks: cellParas }] }],
      });
    } else {
      const p = para(`p${made}`);
      blocks.push(p);
      ids.push(p.id);
      made++;
    }
    group++;
  }
  return { doc: { section: SECTION, blocks }, ids };
}

// ---------------------------------------------------------------------------
// Baseline: the PRE-memoization data-access path, inlined here so the bench has a
// genuine "before" to compare the shipped helpers against. `bodyParagraphs` and
// `bandParagraphs` are NOT memoized, so this rebuilds the whole paragraph
// traversal on every call exactly as the old `paragraphsOf` did, then linear-scans
// — the cost the WeakMap index removed.

const rebuildParagraphsOf = (doc: Document): Paragraph[] => [
  ...bodyParagraphs(doc),
  ...BAND_CONTAINERS.flatMap((band) => bandParagraphs(doc, band)),
  ...Object.values(doc.footnotes ?? {}).flat(),
  ...Object.values(doc.endnotes ?? {}).flat(),
];
const blockByIdBaseline = (doc: Document, id: string): Paragraph | undefined =>
  rebuildParagraphsOf(doc).find((b) => b.id === id);
const blockIndexOfBaseline = (doc: Document, id: string): number =>
  rebuildParagraphsOf(doc).findIndex((b) => b.id === id);

// A scattered, deterministic sample of ids to look up (front, middle, back, …)
// so neither impl is favoured by always hitting index 0.
function sampleIds(ids: string[], k: number): string[] {
  const out: string[] = [];
  const step = Math.max(1, Math.floor(ids.length / k));
  for (let i = 0; i < ids.length && out.length < k; i += step) out.push(ids[i]!);
  return out;
}

// A fresh doc IDENTITY without re-walking content — models one keystroke
// producing a new Document (an edit path-clones the spine; {...doc} shares the
// arrays but gets a new identity, so the WeakMap misses, like a real edit).
const reidentify = (doc: Document): Document => ({ ...doc });

// ---------------------------------------------------------------------------
// Scenarios. LOOKUPS = how many id lookups happen on ONE doc identity before it
// changes. 1 ≈ a lone lookup; 5 ≈ a keystroke through the selection controller
// (textOf + blockIndexOf + the find + neighbour textOf …) on the same doc.

const SIZES = [100, 1_000, 5_000];
const LOOKUPS = [1, 5];

for (const N of SIZES) {
  const { doc: base, ids } = makeDoc(N);

  for (const L of LOOKUPS) {
    const picks = sampleIds(ids, L);

    // Realistic per-keystroke loop: each iteration is a NEW doc identity with L
    // lookups on it. Baseline rebuilds L times; shipped builds once + L map hits.
    describe(`per-keystroke (new doc each iter) — N=${N}, lookups/doc=${L}`, () => {
      bench("baseline blockById (rebuild+scan)", () => {
        const doc = reidentify(base);
        for (const id of picks) {
          blockByIdBaseline(doc, id);
          blockIndexOfBaseline(doc, id);
        }
      });
      bench("shipped  blockById (weakmap O(1))", () => {
        const doc = reidentify(base);
        for (const id of picks) {
          blockById(doc, id);
          blockIndexOf(doc, id);
        }
      });
    });
  }

  // Navigation / stable-doc case: the doc identity does NOT change across many
  // lookups (e.g. holding an arrow key, or many helpers reading the same doc in
  // one render). This is the cache-hit-dominated best case for the index.
  describe(`stable doc (no edits) — N=${N}, 50 scattered lookups`, () => {
    const picks = sampleIds(ids, 50);
    bench("baseline blockById (rebuild+scan)", () => {
      for (const id of picks) blockByIdBaseline(base, id);
    });
    bench("shipped  blockById (weakmap O(1))", () => {
      for (const id of picks) blockById(base, id);
    });
  });
}
