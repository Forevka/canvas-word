// Content-control (OOXML w:sdt) path helpers.
//
// SDT membership is an ORDERED ANCESTRY PATH (outer→inner) so controls can nest:
//   - inline controls put their ids on `CharStyle.sdtPath` (within a paragraph);
//   - block-level controls put their ids on `Block.sdtPath` (whole paragraphs/
//     tables, incl. run-less blocks).
// A run's full enclosing chain is `block.sdtPath ++ run.style.sdtPath`. The props
// for every id live in `Document.sdts`. Paths are treated as immutable (always
// replaced, never mutated in place) and are never stored as `[]` — use undefined.

import type { Block, CharStyle } from "./document";

/** The innermost (most specific) inline control id at a run, or undefined. */
export function innermostSdtId(style: CharStyle): string | undefined {
  const p = style.sdtPath;
  return p && p.length > 0 ? p[p.length - 1] : undefined;
}

/** The innermost block-level control id wrapping a block, or undefined. */
export function blockInnermostSdtId(block: Block): string | undefined {
  const p = block.sdtPath;
  return p && p.length > 0 ? p[p.length - 1] : undefined;
}

/** A run's full enclosing control chain, outer→inner: block ancestry (prefix)
 *  followed by inline ancestry (suffix). Empty array if none. */
export function fullSdtChain(block: Block | undefined, style: CharStyle): string[] {
  const bp = block?.sdtPath ?? [];
  const ip = style.sdtPath ?? [];
  return bp.length === 0 ? ip.slice() : bp.concat(ip);
}

/** Path equality treating `undefined` and `[]` as equal (the empty path). */
export function sdtPathEq(a: string[] | undefined, b: string[] | undefined): boolean {
  const al = a?.length ?? 0;
  const bl = b?.length ?? 0;
  if (al !== bl) return false;
  for (let i = 0; i < al; i++) if (a![i] !== b![i]) return false;
  return true;
}

/** Length of the shared outer prefix of two paths (for LCP grouping on export). */
export function commonPrefixLen(a: string[], b: string[]): number {
  const n = Math.min(a.length, b.length);
  let i = 0;
  while (i < n && a[i] === b[i]) i++;
  return i;
}

/** Whether a run's inline path contains `id` (membership at any nesting level). */
export function inSdt(style: CharStyle, id: string): boolean {
  return style.sdtPath?.includes(id) ?? false;
}

/** Whether a block's path contains `id`. */
export function blockInSdt(block: Block, id: string): boolean {
  return block.sdtPath?.includes(id) ?? false;
}

/** Normalize a candidate path: undefined for empty, else a fresh array. */
export function normSdtPath(path: string[] | undefined): string[] | undefined {
  return path && path.length > 0 ? path.slice() : undefined;
}

/** Append a control id to a path (new array), for wrapping inside existing ones. */
export function pushSdt(path: string[] | undefined, id: string): string[] {
  return path && path.length > 0 ? [...path, id] : [id];
}

/** Remove ONE control id from a path (new array), preserving ancestry; returns
 *  undefined when the path becomes empty. Used by "remove control". */
export function removeSdt(path: string[] | undefined, id: string): string[] | undefined {
  if (!path) return undefined;
  const next = path.filter((x) => x !== id);
  return next.length > 0 ? next : undefined;
}

/** The path prefix up to and INCLUDING `id` — the ancestry that newly inserted
 *  content replacing control `id`'s contents must carry. Returns undefined if
 *  `id` is not on the path. */
export function ancestryThrough(path: string[] | undefined, id: string): string[] | undefined {
  if (!path) return undefined;
  const idx = path.indexOf(id);
  return idx < 0 ? undefined : path.slice(0, idx + 1);
}
