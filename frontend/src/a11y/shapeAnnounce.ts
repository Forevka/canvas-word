// Layer 6 (a11y), shape-scoped. Drawing shapes are painted on the canvas and so
// are invisible to assistive tech (E4 in docs/UX_ANALYSIS.md): selecting, moving,
// or rotating one announced nothing and no shape carried an accessible name. These
// pure helpers turn a ShapeBlock into (a) an accessible NAME — the geometry's
// friendly label plus any text-box text — and (b) the live-region ANNOUNCEMENT
// strings the editor speaks on select / move / rotate. The wiring lives in
// index.ts; the full DOM-selection canvas mirror is still a milestone-6 item.

import type { ShapeBlock, ShapePreset } from "@cw/shared";
import { textOfRuns } from "@cw/shared";

/** Friendly, screen-reader-spoken names for each preset geometry (the OOXML
 *  `a:prstGeom@prst` tokens themselves — "roundRect", "rightArrow" — read poorly). */
const PRESET_LABELS: Record<ShapePreset, string> = {
  rect: "Rectangle",
  roundRect: "Rounded rectangle",
  ellipse: "Ellipse",
  triangle: "Triangle",
  diamond: "Diamond",
  rightArrow: "Right arrow",
  leftArrow: "Left arrow",
  line: "Line",
};

/** Longest text-box excerpt folded into an accessible name before it's truncated —
 *  keeps a paragraph-long text box from producing a sentence-long announcement. */
const NAME_TEXT_CAP = 80;

/** The human label for a shape's geometry: a group reports its member count, a
 *  custom-geometry (freeform) shape is a generic "Shape", otherwise the preset's
 *  friendly name (with an unknown preset falling back to "Shape"). */
export function shapeGeometryLabel(shape: ShapeBlock): string {
  if (shape.group) {
    const n = shape.group.children.length;
    return `Group of ${n} ${n === 1 ? "shape" : "shapes"}`;
  }
  if (shape.geometry.custom) return "Shape";
  return PRESET_LABELS[shape.geometry.preset] ?? "Shape";
}

/** The shape's text-box text flattened to a single trimmed line (paragraphs joined
 *  by spaces); "" when the shape has no body or only whitespace. */
export function shapeText(shape: ShapeBlock): string {
  if (!shape.text) return "";
  return shape.text.blocks
    .map((p) => textOfRuns(p.runs))
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
}

/** A shape's accessible name: the geometry label, plus its text (capped) when it
 *  carries any — e.g. `Rectangle`, `Rectangle: Hello world`, `Group of 3 shapes`.
 *  Used as the selection-frame `aria-label` and as the spine of every announcement. */
export function shapeAccessibleName(shape: ShapeBlock): string {
  const label = shapeGeometryLabel(shape);
  const text = shapeText(shape);
  if (!text) return label;
  const excerpt = text.length > NAME_TEXT_CAP ? `${text.slice(0, NAME_TEXT_CAP)}…` : text;
  return `${label}: ${excerpt}`;
}

/** Live-region message when a shape becomes the selected object. */
export function describeShapeSelected(shape: ShapeBlock): string {
  return `${shapeAccessibleName(shape)}, selected`;
}

/** Live-region message when a shape is moved (anchored-shape drag/nudge commit).
 *  Offsets are the anchor-relative px lever the move writes, rounded for speech. */
export function describeShapeMoved(shape: ShapeBlock, offsetXPx: number, offsetYPx: number): string {
  return `${shapeGeometryLabel(shape)} moved to ${Math.round(offsetXPx)}, ${Math.round(offsetYPx)} pixels`;
}

/** Live-region message when a shape's rotation is committed; `deg` is the new
 *  normalized rotation ([0,360)), where 0 means the rotation was cleared. */
export function describeShapeRotated(shape: ShapeBlock, deg: number): string {
  const label = shapeGeometryLabel(shape);
  return deg === 0 ? `${label}, rotation removed` : `${label} rotated to ${Math.round(deg)} degrees`;
}
