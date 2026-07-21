// Pure placement math for a floating toolbar anchored to a selection/object rect.
// DOM-free so it's unit-testable; the same flip-above/clamp logic the image
// mini-toolbar uses inline (editorApp.ts), extracted for the text selection bar.

/** A viewport (client-px) rectangle to anchor against. */
export interface AnchorRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface BarSize {
  width: number;
  height: number;
}

export interface Viewport {
  width: number;
  height: number;
  /** Top edge of the usable area in client px (default 0). Set this to the editor
   *  content area's top — i.e. below a fixed ribbon — so bars neither render over
   *  the ribbon nor linger there once the anchor scrolls up behind it. */
  top?: number;
}

export interface Placement {
  left: number;
  top: number;
}

/** Where to render a `bar` of the given size so it sits centered above `anchor`,
 *  flipping below when there isn't room above `topGuard`, and clamped `margin` px
 *  inside the viewport. Coordinates are integer client px for a `position:fixed`
 *  element. */
export function placeSelectionBar(
  anchor: AnchorRect,
  bar: BarSize,
  viewport: Viewport,
  opts: { gap?: number; margin?: number; topGuard?: number } = {},
): Placement {
  const gap = opts.gap ?? 8;
  const margin = opts.margin ?? 8;
  // Keep clear of the content-area top (below a fixed ribbon): flipping above must
  // not collide with it, and the placed bar must never sit above it.
  const topGuard = Math.max(opts.topGuard ?? 56, viewport.top ?? 0);

  const centered = anchor.left + anchor.width / 2 - bar.width / 2;
  const maxLeft = Math.max(margin, viewport.width - bar.width - margin);
  const left = Math.min(Math.max(margin, centered), maxLeft);

  const above = anchor.top - bar.height - gap;
  const below = anchor.top + anchor.height + gap;
  // Flip below when there isn't room above the guard; clamp so we never draw above it.
  const top = Math.max(above < topGuard ? below : above, viewport.top ?? 0);

  return { left: Math.round(left), top: Math.round(top) };
}

/** Whether the anchor is at all within the (optionally padded) viewport — used to
 *  hide the bar once the selection scrolls out of view. The top edge honours
 *  `viewport.top`, so an anchor that has scrolled up behind the ribbon counts as
 *  out of view (the bar hides instead of hovering over the ribbon). */
export function anchorInView(anchor: AnchorRect, viewport: Viewport, pad = 0): boolean {
  return (
    anchor.top + anchor.height >= (viewport.top ?? 0) - pad &&
    anchor.top <= viewport.height + pad &&
    anchor.left + anchor.width >= -pad &&
    anchor.left <= viewport.width + pad
  );
}
