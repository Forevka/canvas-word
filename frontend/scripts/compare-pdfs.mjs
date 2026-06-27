// Byte-compare two directories of PDFs (Node-host vs ClearScript-host output of the
// SAME bundle). Reports identical / size diff / first differing offset / total
// differing bytes, and the surrounding bytes so we can see WHAT differs.
// Usage: node compare-pdfs.mjs <dirA> <dirB>
import { readFileSync, readdirSync } from "node:fs";

const [dirA, dirB] = process.argv.slice(2);
if (!dirA || !dirB) { console.error("usage: compare-pdfs.mjs <dirA> <dirB>"); process.exit(2); }

const ctx = (buf, off, n = 60) => {
  const s = Math.max(0, off - 8), e = Math.min(buf.length, off + n);
  return JSON.stringify(buf.slice(s, e).toString("latin1"));
};

let allIdentical = true;
const files = new Set([
  ...readdirSync(dirA).filter((f) => f.endsWith(".pdf")),
  ...readdirSync(dirB).filter((f) => f.endsWith(".pdf")),
]);
for (const f of [...files].sort()) {
  let a, b;
  try { a = readFileSync(`${dirA}/${f}`); b = readFileSync(`${dirB}/${f}`); }
  catch { console.log(`MISSING pair for ${f}`); allIdentical = false; continue; }

  if (a.equals(b)) { console.log(`IDENTICAL  ${f}  (${a.length} bytes)`); continue; }

  allIdentical = false;
  const min = Math.min(a.length, b.length);
  let first = -1, diffs = 0;
  for (let i = 0; i < min; i++) if (a[i] !== b[i]) { if (first < 0) first = i; diffs++; }
  diffs += Math.abs(a.length - b.length);

  // Collect the distinct differing regions (runs of differing bytes).
  const regions = [];
  let i = 0;
  while (i < min) {
    if (a[i] !== b[i]) {
      const start = i;
      while (i < min && a[i] !== b[i]) i++;
      regions.push([start, i]);
    } else i++;
  }
  if (a.length !== b.length) {
    if (first < 0) first = min; // shared prefix, only the tail differs
    regions.push([min, Math.max(a.length, b.length)]);
  }
  console.log(`DIFFERS    ${f}  sizes=${a.length}/${b.length}  firstDiff@${first}  diffBytes=${diffs}  regions=${regions.length}`);
  for (const [s, e] of regions.slice(0, 6)) {
    console.log(`   @${s}..${e}\n     A=${ctx(a, s, e - s + 20)}\n     B=${ctx(b, s, e - s + 20)}`);
  }
  if (regions.length > 6) console.log(`   …and ${regions.length - 6} more regions`);
}
console.log(allIdentical ? "\nALL IDENTICAL" : "\nSOME DIFFER (see above)");
