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
  // Keep clear of the ribbon area near the top when flipping above would collide.
  const topGuard = opts.topGuard ?? 56;

  const centered = anchor.left + anchor.width / 2 - bar.width / 2;
  const maxLeft = Math.max(margin, viewport.width - bar.width - margin);
  const left = Math.min(Math.max(margin, centered), maxLeft);

  const above = anchor.top - bar.height - gap;
  const below = anchor.top + anchor.height + gap;
  const top = above < topGuard ? below : above;

  return { left: Math.round(left), top: Math.round(top) };
}

/** Whether the anchor is at all within the (optionally padded) viewport — used to
 *  hide the bar once the selection scrolls out of view. */
export function anchorInView(anchor: AnchorRect, viewport: Viewport, pad = 0): boolean {
  return (
    anchor.top + anchor.height >= -pad &&
    anchor.top <= viewport.height + pad &&
    anchor.left + anchor.width >= -pad &&
    anchor.left <= viewport.width + pad
  );
}
