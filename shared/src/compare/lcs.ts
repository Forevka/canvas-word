// LCS-based sequence aligner — the shared diff primitive for document compare.
// Aligns two sequences by a string key and returns coalesced hunks (equal /
// delete / insert / replace). Used at the BLOCK level (key = blockKey) and the
// TOKEN level (key = token text) by diffDocuments. Pure + DOM-free.
//
// Classic O(n·m) LCS dynamic program: fine for the block counts and paragraph
// token counts real documents have. A Myers refinement is a later optimization
// if a pathological paragraph shows up; the hunk shape it produces is identical.

export type DiffTag = "equal" | "delete" | "insert" | "replace";

/** One aligned hunk. `a`/`b` are the covered items from each side (one is empty
 *  for a pure insert/delete); `ai`/`bi` are the hunk's start indices in A/B. */
export interface AlignOp<T> {
  tag: DiffTag;
  a: T[];
  b: T[];
  ai: number;
  bi: number;
}

/** Per-element decision produced by the LCS backtrack, before run-coalescing. */
type Step = { kind: "equal" | "delete" | "insert" };

/** Align `a` and `b` by `key`, returning coalesced hunks in sequence order. A
 *  maximal delete run immediately followed by a maximal insert run (or vice
 *  versa) is merged into a single `replace` hunk. */
export function align<T>(a: readonly T[], b: readonly T[], key: (t: T) => string): AlignOp<T>[] {
  const n = a.length;
  const m = b.length;
  const ka = a.map(key);
  const kb = b.map(key);

  // dp[i][j] = LCS length of a[i..] and b[j..].
  const dp: number[][] = Array.from({ length: n + 1 }, () => new Array<number>(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      dp[i]![j] = ka[i] === kb[j] ? dp[i + 1]![j + 1]! + 1 : Math.max(dp[i + 1]![j]!, dp[i]![j + 1]!);
    }
  }

  // Backtrack into a flat list of per-element steps.
  const steps: Step[] = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (ka[i] === kb[j]) {
      steps.push({ kind: "equal" });
      i++;
      j++;
    } else if (dp[i + 1]![j]! >= dp[i]![j + 1]!) {
      steps.push({ kind: "delete" });
      i++;
    } else {
      steps.push({ kind: "insert" });
      j++;
    }
  }
  while (i < n) {
    steps.push({ kind: "delete" });
    i++;
  }
  while (j < m) {
    steps.push({ kind: "insert" });
    j++;
  }

  // Coalesce consecutive same-kind steps into runs, then fuse an adjacent
  // delete-run + insert-run (either order) into one `replace` hunk.
  const out: AlignOp<T>[] = [];
  let ai = 0;
  let bi = 0;
  let s = 0;
  while (s < steps.length) {
    const kind = steps[s]!.kind;
    let e = s;
    while (e < steps.length && steps[e]!.kind === kind) e++;
    const count = e - s;
    if (kind === "equal") {
      out.push({ tag: "equal", a: a.slice(ai, ai + count), b: b.slice(bi, bi + count), ai, bi });
      ai += count;
      bi += count;
    } else if (kind === "delete") {
      out.push({ tag: "delete", a: a.slice(ai, ai + count), b: [], ai, bi });
      ai += count;
    } else {
      out.push({ tag: "insert", a: [], b: b.slice(bi, bi + count), ai, bi });
      bi += count;
    }
    s = e;
  }

  return coalesceReplace(out);
}

/** Fuse a delete hunk adjacent to an insert hunk (either order) into a replace. */
function coalesceReplace<T>(hunks: AlignOp<T>[]): AlignOp<T>[] {
  const out: AlignOp<T>[] = [];
  for (const h of hunks) {
    const prev = out[out.length - 1];
    if (
      prev &&
      ((prev.tag === "delete" && h.tag === "insert") || (prev.tag === "insert" && h.tag === "delete"))
    ) {
      const del = prev.tag === "delete" ? prev : h;
      const ins = prev.tag === "insert" ? prev : h;
      out[out.length - 1] = {
        tag: "replace",
        a: del.a,
        b: ins.b,
        ai: Math.min(prev.ai, h.ai),
        bi: Math.min(prev.bi, h.bi),
      };
    } else {
      out.push(h);
    }
  }
  return out;
}
