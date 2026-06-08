// Post-import Table-of-Contents recognition. A Word TOC arrives as paragraphs
// styled "TOC 1".."TOC 9", each holding the heading text, a tab to a right tab
// stop with a dot leader, and a PAGEREF page number — all hyperlinked to the
// heading's bookmark. We lose the tab/leader on import (tabs → spaces) and the
// page number becomes stale literal text.
//
// This pass turns those entries into the editor's native TOC mechanism: it
// resolves each entry's anchor to its heading via Document.bookmarks, strips
// the literal tab + page number, and sets ParaStyle.tocEntry — so the layout
// engine renders a real dot leader and a LIVE page number (recomputed from the
// current pagination), exactly like Word's "update field".

import type { Document, Run } from "./document";

export function recognizeImportedToc(doc: Document): Document {
  const sheet = doc.stylesheet;
  if (!sheet) return doc;

  // styleId → TOC level, from the style's display name ("toc 1" … "TOC 9").
  const tocLevel = new Map<string, number>();
  for (const s of sheet.styles) {
    const m = s.name.match(/^toc\s*(\d)\b/i);
    if (m) tocLevel.set(s.id, Number(m[1]));
  }
  if (tocLevel.size === 0) return doc;

  let changed = false;
  const blocks = doc.blocks.map((b) => {
    if (b.kind !== "paragraph" || b.style.tocEntry) return b;
    const styleId = b.style.namedStyle;
    const level = styleId ? tocLevel.get(styleId) : undefined;
    if (level === undefined) return b;

    // The heading bookmark is the entry's in-document anchor link (#name).
    const anchorRun = b.runs.find((r) => r.style.link?.startsWith("#"));
    const anchor = anchorRun?.style.link?.slice(1);
    const targetId = anchor ? doc.bookmarks?.[anchor] : undefined;
    if (!targetId) return b; // unresolved — leave the entry as imported

    const runs = stripTocTail(b.runs);
    if (runs.length === 0) return b; // nothing but the page number? skip
    changed = true;
    return {
      ...b,
      revision: b.revision + 1,
      runs,
      style: { ...b.style, tocEntry: { targetId, level } },
    };
  });
  return changed ? { ...doc, blocks } : doc;
}

/** Drop the trailing tab-gap + page number ("Heading␣␣␣␣12" → "Heading"). The
 *  import converted the leader tab to spaces, so the tail is 2+ whitespace/dot
 *  chars followed by digits at the very end. */
function stripTocTail(runs: Run[]): Run[] {
  const full = runs.map((r) => r.text).join("");
  const tail = full.match(/[\s.…]{2,}\d+[\s.…]*$/);
  const cut = tail ? full.length - tail[0].length : full.length;

  const out: Run[] = [];
  let acc = 0;
  for (const r of runs) {
    const end = acc + r.text.length;
    if (end <= cut) out.push(r);
    else if (acc < cut) out.push({ ...r, text: r.text.slice(0, cut - acc) });
    acc = end;
  }
  // Trim trailing whitespace/leader left on the final kept run.
  if (out.length > 0) {
    const last = out[out.length - 1]!;
    const trimmed = last.text.replace(/[\s.…]+$/, "");
    if (trimmed.length === 0) out.pop();
    else out[out.length - 1] = { ...last, text: trimmed };
  }
  return out;
}
