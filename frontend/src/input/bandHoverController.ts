// Hover affordance for the header/footer margin bands. When the pointer rests
// over a header or footer region — and the editor is NOT already in band-edit
// ("story") mode — this overlays, purely in the DOM (like the object-selection
// frame; no canvas repaint, no model op):
//   • a dashed outline delineating the band "area", with a soft tint and a small
//     corner label ("Header" / "Footer"), so the band is visually distinct from
//     the body content, and
//   • an "Edit" button the user can click to enter band-edit mode — a
//     discoverable alternative to the existing double-click, which is preserved.
// The overlay lives in the page placeholder element, positioned in zoomed page
// pixels (sizes are document px), matching the object-frame overlay pattern.

import type { Rect } from "../layout/geometry";

export type BandName = "header" | "footer";

/** The slice of page geometry the band region is derived from — a structural
 *  subset of the layout `Page`, so a `Page` can be passed straight in. */
export interface BandPageGeom {
  index: number;
  widthPx: number;
  heightPx: number;
  marginPx: { left: number; right: number };
  contentTopPx: number;
  contentBottomPx: number;
}

/** The band's clickable/paintable region on its page (document px). The header
 *  occupies everything above the body's real content top; the footer everything
 *  below the content bottom (tall bands push the content box, so these track it).
 *  Horizontally the box spans the content width (inside the L/R margins) — where
 *  header/footer text actually lives — which reads as the band area, Word-style. */
export function bandRegionRect(pg: BandPageGeom, band: BandName): Rect {
  const contentWidth = pg.widthPx - pg.marginPx.left - pg.marginPx.right;
  return band === "header"
    ? { pageIndex: pg.index, x: pg.marginPx.left, y: 0, width: contentWidth, height: pg.contentTopPx }
    : {
        pageIndex: pg.index,
        x: pg.marginPx.left,
        y: pg.contentBottomPx,
        width: contentWidth,
        height: pg.heightPx - pg.contentBottomPx,
      };
}

export interface BandHoverDeps {
  getPageElement(pageIndex: number): HTMLElement | null;
  /** Presentational zoom — the overlay lives in zoomed page px, sizes are doc px. */
  getZoom(): number;
  /** Enter band-edit (story) mode for the clicked band/page. */
  onEdit(band: BandName, pageIndex: number): void;
}

export interface BandHoverController {
  /** Show the affordance over `rect` (band region, doc px) on `pageIndex`. */
  show(band: BandName, pageIndex: number, rect: Rect): void;
  hide(): void;
  destroy(): void;
}

const ACCENT = "#1a73e8";

export function createBandHoverController(deps: BandHoverDeps): BandHoverController {
  // The outline box (the "area" lines) — non-interactive so it never blocks a
  // click/double-click reaching the canvas underneath (caret placement, or the
  // preserved double-click-to-edit). Only the button opts back into pointer events.
  const root = document.createElement("div");
  root.className = "cw-band-hover";
  root.style.cssText =
    "position:absolute;display:none;box-sizing:border-box;pointer-events:none;z-index:2;" +
    `border:1px dashed ${ACCENT};background:rgba(26,115,232,0.04);`;

  // Corner label pill identifying the band ("Header" / "Footer").
  const label = document.createElement("div");
  label.className = "cw-band-hover-label";
  label.style.cssText =
    "position:absolute;left:-1px;padding:1px 6px;font:600 11px/1.45 system-ui,-apple-system,sans-serif;" +
    `color:#fff;background:${ACCENT};pointer-events:none;white-space:nowrap;user-select:none;`;
  root.appendChild(label);

  // The clickable "Edit" affordance. Kept inside the band region (never spilling
  // into the body) so that hovering it still resolves to the same band and the
  // overlay stays put.
  const button = document.createElement("button");
  button.type = "button";
  button.className = "cw-band-hover-edit";
  button.textContent = "✎ Edit";
  button.style.cssText =
    "position:absolute;right:6px;padding:3px 9px;font:600 11px/1 system-ui,-apple-system,sans-serif;" +
    `color:#fff;background:${ACCENT};border:none;border-radius:4px;cursor:pointer;pointer-events:auto;` +
    "box-shadow:0 1px 3px rgba(0,0,0,0.25);white-space:nowrap;";
  root.appendChild(button);

  let current: { band: BandName; pageIndex: number } | null = null;
  let lastKey = "";

  // Swallow the press so it doesn't fall through to the canvas mousedown
  // (selection/caret routing) beneath the overlay; the click drives the action.
  const onButtonDown = (ev: Event): void => {
    ev.preventDefault();
    ev.stopPropagation();
  };
  const onButtonClick = (ev: MouseEvent): void => {
    ev.preventDefault();
    ev.stopPropagation();
    if (current) deps.onEdit(current.band, current.pageIndex);
  };
  button.addEventListener("mousedown", onButtonDown);
  button.addEventListener("pointerdown", onButtonDown);
  button.addEventListener("click", onButtonClick);

  const hide = (): void => {
    if (root.style.display === "none") return;
    root.style.display = "none";
    current = null;
    lastKey = "";
  };

  return {
    show(band: BandName, pageIndex: number, rect: Rect): void {
      const host = deps.getPageElement(pageIndex);
      if (!host) {
        hide();
        return;
      }
      const z = deps.getZoom();
      // Skip redundant DOM writes while the pointer glides within the same band
      // (this runs every hover frame) — avoids per-frame layout thrash / flicker.
      const key = `${band}:${pageIndex}:${rect.x}:${rect.y}:${rect.width}:${rect.height}:${z}`;
      if (key === lastKey && root.parentElement === host) return;
      lastKey = key;
      current = { band, pageIndex };
      if (root.parentElement !== host) host.appendChild(root);

      root.style.display = "block";
      root.style.left = `${rect.x * z}px`;
      root.style.top = `${rect.y * z}px`;
      root.style.width = `${rect.width * z}px`;
      root.style.height = `${rect.height * z}px`;

      const isHeader = band === "header";
      label.textContent = isHeader ? "Header" : "Footer";
      // Pin the label to the OUTER edge (header top, footer bottom) and the button
      // to the INNER edge (by the body boundary) — Word-like, and it keeps the
      // button within the band whichever way the band is oriented.
      label.style.top = isHeader ? "-1px" : "auto";
      label.style.bottom = isHeader ? "auto" : "-1px";
      label.style.borderRadius = isHeader ? "0 0 3px 0" : "0 3px 0 0";
      button.style.top = isHeader ? "auto" : "6px";
      button.style.bottom = isHeader ? "6px" : "auto";
    },
    hide,
    destroy(): void {
      button.removeEventListener("mousedown", onButtonDown);
      button.removeEventListener("pointerdown", onButtonDown);
      button.removeEventListener("click", onButtonClick);
      root.remove();
    },
  };
}
