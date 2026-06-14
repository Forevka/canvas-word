// Review layer → paint decorations. Walks the review overlay against the layout
// tree once and emits page-local colored boxes (insertion underline, deletion
// strike, comment highlight) + margin change-bars and comment pins. Reuses
// geometry.selectionRects (the same range→rects helper selection/search use), so
// it never measures text and never triggers relayout — a pure repaint feed.

import { colorForId, type ReviewLayer } from "@cw/shared";
import type { LayoutTree } from "../layout/layoutTree";
import { caretRect, selectionRects, type GeoScope, type Rect } from "../layout/geometry";
import type { ReviewDecoBox, ReviewDecorations, ReviewPin } from "../paint/renderer";

const collapsed = (a: { blockId: string; offset: number }, b: { blockId: string; offset: number }): boolean =>
  a.blockId === b.blockId && a.offset === b.offset;

export function decorate(review: ReviewLayer, tree: LayoutTree, scope?: GeoScope): ReviewDecorations {
  const comments: ReviewDecoBox[] = [];
  const inserts: ReviewDecoBox[] = [];
  const deletes: ReviewDecoBox[] = [];
  const changeBars: ReviewDecoBox[] = [];
  const pins: ReviewPin[] = [];

  // A change-bar in the left margin spanning the lines a suggestion touches.
  const addChangeBar = (rects: Rect[], color: string): void => {
    const byPage = new Map<number, { top: number; bot: number }>();
    for (const r of rects) {
      const cur = byPage.get(r.pageIndex);
      if (cur) {
        cur.top = Math.min(cur.top, r.y);
        cur.bot = Math.max(cur.bot, r.y + r.height);
      } else byPage.set(r.pageIndex, { top: r.y, bot: r.y + r.height });
    }
    for (const [pageIndex, { top, bot }] of byPage) {
      const page = tree.pages[pageIndex];
      if (!page) continue;
      changeBars.push({ pageIndex, x: page.marginPx.left - 9, y: top, width: 2.5, height: bot - top, color });
    }
  };

  for (const s of review.suggestions) {
    if (collapsed(s.anchor.start, s.anchor.end)) continue;
    const color = colorForId(s.author.id);
    const rects = selectionRects(tree, { anchor: s.anchor.start, focus: s.anchor.end }, scope);
    if (rects.length === 0) continue;
    const boxes = rects.map((r) => ({ ...r, color }));
    if (s.kind === "insert") inserts.push(...boxes);
    else if (s.kind === "delete") deletes.push(...boxes);
    addChangeBar(rects, color); // every tracked change gets a change bar
  }

  for (const t of review.threads) {
    const color = colorForId(t.comments[0]?.author.id ?? "thread");
    const point = collapsed(t.anchor.start, t.anchor.end);
    let rects: Rect[];
    if (point) {
      const c = caretRect(tree, t.anchor.start, scope);
      rects = c ? [{ pageIndex: c.pageIndex, x: c.x, y: c.y, width: 2, height: c.height }] : [];
    } else {
      rects = selectionRects(tree, { anchor: t.anchor.start, focus: t.anchor.end }, scope);
      for (const r of rects) comments.push({ ...r, color });
    }
    if (rects.length === 0) continue;
    const first = rects[0]!;
    const page = tree.pages[first.pageIndex];
    if (page) {
      pins.push({
        pageIndex: first.pageIndex,
        x: page.widthPx - page.marginPx.right + 12,
        y: first.y,
        color,
        threadId: t.id,
        resolved: t.status === "resolved",
      });
    }
  }

  return { comments, inserts, deletes, changeBars, pins };
}
