// Round-trip fidelity test for nested outer->inner content controls.
// import(docx) -> inspect doc.sdts -> export(docx) -> re-import -> compare.
//
// Run from repo root:  node test-fixtures/scrape/roundtrip-test.mjs
import { readFileSync } from 'node:fs';
import { runImport } from '../../frontend/dist-node/import.js';
import { runExport } from '../../frontend/dist-node/export.js';

const path = new URL('./scrape-sample.docx', import.meta.url);
const bytes = new Uint8Array(readFileSync(path));

function summarizeSdts(doc, label) {
  console.log(`\n=== ${label} ===`);
  const sdts = doc.sdts ?? {};
  console.log('doc.sdts count:', Object.keys(sdts).length);
  for (const [id, props] of Object.entries(sdts)) {
    let tag = props.tag;
    let parsed = null;
    try { parsed = tag ? JSON.parse(tag) : null; } catch {}
    const kind = parsed?.k === 'fb' ? `FB[target=${parsed.target},orig=${JSON.stringify(parsed.orig)}]`
               : (parsed?.Name ? `SECTION[${parsed.Name}]` : 'OTHER');
    console.log(`  ${id}: alias=${JSON.stringify(props.alias)} -> ${kind}`);
  }
  // run-level membership: full enclosing chain = block.sdtPath ++ run.style.sdtPath
  const runLog = [];
  const walkBlocks = (blocks, blockPath = []) => {
    for (const b of blocks ?? []) {
      const bp = b.sdtPath ?? blockPath;
      if (b.kind === 'paragraph') {
        for (const r of b.runs ?? []) {
          const chain = [...bp, ...(r.style?.sdtPath ?? [])];
          if (r.text?.trim()) runLog.push(`    "${r.text}" -> [${chain.join(' > ')}]`);
        }
      } else if (b.kind === 'table') {
        for (const row of b.rows ?? []) for (const cell of row.cells ?? []) walkBlocks(cell.blocks, bp);
      }
    }
  };
  walkBlocks(doc.blocks);
  console.log('  per-run enclosing control chain (outer > inner):');
  console.log(runLog.join('\n'));
}

const imp1 = runImport(bytes);
const doc1 = imp1.doc ?? imp1;
summarizeSdts(doc1, 'AFTER FIRST IMPORT');
if (imp1.warnings?.length) console.log('import warnings:', imp1.warnings);

// export back to docx (runExport is async)
const exp = await runExport(doc1, 'docx', imp1.images ?? imp1.media ?? undefined);
const outBytes = exp instanceof Uint8Array ? exp : (exp?.bytes ?? exp?.data ?? exp?.docx ?? exp);
console.log('\nexported docx bytes:', outBytes?.length ?? '(unknown shape)');

const imp2 = runImport(new Uint8Array(outBytes));
const doc2 = imp2.doc ?? imp2;
summarizeSdts(doc2, 'AFTER EXPORT + RE-IMPORT (unedited)');
if (imp2.warnings?.length) console.log('re-import warnings:', imp2.warnings);

// --- EDIT leg: change the value inside the "Appraisal Fee" inner control -----
// Find the run carrying sdt2 (Appraisal Fee) and rewrite its text, as an edit would.
const editTargetSdt = Object.entries(doc1.sdts).find(([,p]) => p.alias === 'Appraisal Fee')?.[0];
let edited = 0;
const editRuns = (blocks) => {
  for (const b of blocks ?? []) {
    if (b.kind === 'paragraph') {
      for (const r of b.runs ?? []) if (r.style?.sdtPath?.includes(editTargetSdt)) { r.text = '$300.00'; edited++; }
    } else if (b.kind === 'table') {
      for (const row of b.rows ?? []) for (const cell of row.cells ?? []) editRuns(cell.blocks);
    }
  }
};
editRuns(doc1.blocks);
console.log(`\n=== EDIT: rewrote ${edited} run(s) in ${editTargetSdt} ($200.00 -> $300.00) ===`);
const exp2 = await runExport(doc1, 'docx', imp1.images ?? undefined);
const outBytes2 = exp2 instanceof Uint8Array ? exp2 : (exp2?.bytes ?? exp2?.docx ?? exp2);
const imp3 = runImport(new Uint8Array(outBytes2));
const doc3 = imp3.doc ?? imp3;
// report the Appraisal Fee binding's tag + current value after edit round-trip
const feeId = Object.entries(doc3.sdts).find(([,p]) => p.alias === 'Appraisal Fee')?.[0];
let feeText = '';
const collect = (blocks) => { for (const b of blocks ?? []) {
  if (b.kind === 'paragraph') for (const r of b.runs ?? []) { if (r.style?.sdtPath?.includes(feeId)) feeText += r.text; }
  else if (b.kind === 'table') for (const row of b.rows ?? []) for (const cell of row.cells ?? []) collect(cell.blocks);
}};
collect(doc3.blocks);
let feeTag = null; try { feeTag = JSON.parse(doc3.sdts[feeId].tag); } catch {}
console.log('AFTER EDIT round-trip — Appraisal Fee control:');
console.log('  tag orig (unchanged baseline):', JSON.stringify(feeTag?.orig));
console.log('  current value text:', JSON.stringify(feeText));
console.log('  => scrape would report:', feeTag ? `${feeTag.target}/${feeTag.fieldName}: ${feeTag.orig} -> ${feeText}` : '(tag lost!)');
