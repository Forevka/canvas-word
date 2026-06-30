// Step-1 measurement: does a WeakMap-memoized id index actually beat the current
// rebuild-then-scan data-access layer, and how does the win scale with document
// size + lookups-per-doc?
//
// Baseline = the REAL production helpers from @cw/shared (paragraphsOf/blockById/
// blockIndexOf), each of which rebuilds the full paragraph traversal then linear-
// scans. Candidate = the same traversal memoized on document identity (a
// WeakMap<Document, …>), giving O(1) Map lookups after a one-time build per doc.
//
// Run:  npx vitest bench src/perf/dataAccess.bench.ts   (from frontend/)

import { bench, describe } from "vitest";
import {
  blockById,
  blockIndexOf,
  makeDefaultCharStyle,
  makeDefaultParaStyle,
  paragraphsOf,
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
// Candidate: WeakMap-memoized index keyed on Document identity. Because the model
// is immutable (every edit returns a new Document object), a new doc simply misses
// the cache; the old doc is GC'd. The traversal itself reuses the production
// paragraphsOf (done once per doc), so this isolates the win of "build once, then
// O(1)" vs "rebuild + scan every call".

const indexCache = new WeakMap<Document, { list: Paragraph[]; byId: Map<string, number> }>();
function docIndex(doc: Document) {
  let e = indexCache.get(doc);
  if (!e) {
    const list = paragraphsOf(doc);
    const byId = new Map<string, number>();
    for (let i = 0; i < list.length; i++) byId.set(list[i]!.id, i);
    e = { list, byId };
    indexCache.set(doc, e);
  }
  return e;
}
const blockByIdNew = (doc: Document, id: string): Paragraph | undefined => {
  const e = docIndex(doc);
  const i = e.byId.get(id);
  return i === undefined ? undefined : e.list[i];
};
const blockIndexOfNew = (doc: Document, id: string): number => docIndex(doc).byId.get(id) ?? -1;

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
    // lookups on it. Baseline rebuilds L times; candidate builds once + L map hits.
    describe(`per-keystroke (new doc each iter) — N=${N}, lookups/doc=${L}`, () => {
      bench("current  blockById (rebuild+scan)", () => {
        const doc = reidentify(base);
        for (const id of picks) {
          blockById(doc, id);
          blockIndexOf(doc, id);
        }
      });
      bench("indexed  blockById (weakmap O(1))", () => {
        const doc = reidentify(base);
        for (const id of picks) {
          blockByIdNew(doc, id);
          blockIndexOfNew(doc, id);
        }
      });
    });
  }

  // Navigation / stable-doc case: the doc identity does NOT change across many
  // lookups (e.g. holding an arrow key, or many helpers reading the same doc in
  // one render). This is the cache-hit-dominated best case for the index.
  describe(`stable doc (no edits) — N=${N}, 50 scattered lookups`, () => {
    const picks = sampleIds(ids, 50);
    bench("current  blockById (rebuild+scan)", () => {
      for (const id of picks) blockById(base, id);
    });
    bench("indexed  blockById (weakmap O(1))", () => {
      for (const id of picks) blockByIdNew(base, id);
    });
  });
}
