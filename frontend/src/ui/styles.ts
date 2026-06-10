// Component CSS for the editor chrome (ribbon, ruler, outline, status bar,
// popovers). Injected once by WordCanvas so the package is self-contained — the
// host page needs no stylesheet. The only change from the original index.html
// <style> is that the page-level `html, body` rules become a `.wordcanvas-root`
// container rule, so embedding doesn't restyle the host's body.
//
// NOTE (v1): structural selectors are ids (#toolbar/#app/#ruler/#outline/
// #statusbar). One WordCanvas per page is supported; multi-instance / host-id
// isolation (shadow DOM or class-scoped selectors) is a follow-up.

const STYLE_ID = "wordcanvas-styles";

const CSS = `
.wordcanvas-root {
  display: flex; flex-direction: column; height: 100%; min-height: 0; overflow-x: hidden;
  font-family: "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
  color: #323130;
}

/* ===== Word-style ribbon ============================================ */
#toolbar {
  flex: 0 0 auto; display: flex; flex-direction: column;
  background: #f3f2f1; border-bottom: 1px solid #e1dfdd; user-select: none;
}

/* --- tab strip --- */
.rib-tabs {
  display: flex; align-items: flex-end; gap: 1px; height: 32px;
  padding: 4px 8px 0; background: #f3f2f1;
}
.rib-tab {
  border: none; background: transparent; cursor: pointer;
  font: inherit; font-size: 13px; color: #323130;
  padding: 5px 11px 6px; border-radius: 4px 4px 0 0; position: relative;
}
.rib-tab:hover { background: #eceae9; }
.rib-tab.active {
  background: #fff; color: #2b579a; font-weight: 600;
  box-shadow: 0 -1px 2px rgba(0,0,0,0.05);
}
.rib-tab.active::after {
  content: ""; position: absolute; left: 8px; right: 8px; bottom: 0;
  height: 2px; background: #2b579a;
}
.rib-tab.file { background: #2b579a; color: #fff; font-weight: 600; }
.rib-tab.file:hover { background: #21457e; }

/* --- ribbon body: one panel visible at a time --- */
.rib-bodies {
  background: #fff; border-top: 1px solid #e1dfdd; border-bottom: 1px solid #e1dfdd;
}
.rib-panel { display: none; align-items: stretch; min-height: 94px; padding: 0 2px; }
.rib-panel.active { display: flex; }

/* --- group: stacked controls + caption, divider on the right --- */
.rib-group {
  display: flex; flex-direction: column; align-items: center;
  padding: 5px 7px 3px; position: relative;
}
.rib-group + .rib-group::before {
  content: ""; position: absolute; left: 0; top: 8px; bottom: 8px;
  width: 1px; background: #e1dfdd;
}
.rib-controls { display: flex; align-items: center; gap: 2px; flex: 1; }
/* two stacked rows (Font, Paragraph) */
.rib-rows { display: flex; flex-direction: column; gap: 3px; flex: 1; justify-content: center; }
.rib-row { display: flex; align-items: center; gap: 2px; }
.rib-label {
  font-size: 11px; color: #605e5c; padding-top: 3px; white-space: nowrap;
}

/* --- buttons --- */
#toolbar button.rib-btn {
  min-width: 26px; height: 26px; display: inline-flex; align-items: center; justify-content: center;
  border: 1px solid transparent; border-radius: 4px; background: transparent; cursor: pointer;
  color: #323130; font-size: 13px; padding: 0 4px; gap: 1px;
}
#toolbar button.rib-btn:hover { background: #e1dfdd; }
#toolbar button.rib-btn:active { background: #d2d0ce; }
#toolbar button.rib-btn.active { background: #cfe3fb; color: #0b57d0; border-color: #b3d3f5; }
#toolbar button.rib-btn.active:hover { background: #bcd8fa; }
#toolbar button.rib-btn.active svg { color: #0b57d0; }
#toolbar button.rib-btn svg { width: 16px; height: 16px; display: block; }
#toolbar button.rib-btn .caret { width: 8px; height: 8px; }

/* large button (Paste): icon over caption */
#toolbar button.rib-big {
  flex-direction: column; height: 100%; min-width: 48px; padding: 4px 6px; gap: 2px; justify-content: center;
}
#toolbar button.rib-big svg { width: 26px; height: 26px; }
#toolbar button.rib-big .big-cap { font-size: 11px; display: flex; align-items: center; gap: 1px; }

/* swatch button: icon row + colour underline */
#toolbar button.rib-swatch { flex-direction: column; gap: 0; padding: 2px 3px; }
#toolbar button.rib-swatch .row { display: flex; align-items: center; gap: 1px; }
#toolbar button.rib-swatch .bar { width: 18px; height: 3px; margin-top: 1px; border-radius: 1px; }

/* disabled stub (feature not supported by the engine yet) */
#toolbar button.rib-btn:disabled { opacity: 0.38; cursor: default; }
#toolbar button.rib-btn:disabled:hover { background: transparent; }

/* form controls */
#toolbar select, #toolbar input[type="number"] {
  height: 24px; border: 1px solid #c8c6c4; border-radius: 3px; background: #fff;
  font: inherit; font-size: 13px; color: #323130;
}
#toolbar select:hover, #toolbar input[type="number"]:hover { border-color: #8a8886; }

/* styles gallery */
.rib-gallery {
  display: flex; align-items: center; gap: 4px; height: 64px; padding: 0 4px;
  border: 1px solid #c8c6c4; border-radius: 4px; background: #fff;
  overflow-x: auto; max-width: 340px;
}
.rib-gallery::-webkit-scrollbar { height: 8px; }
.rib-gallery::-webkit-scrollbar-thumb { background: #c1c1c1; border-radius: 4px; }
.style-card {
  flex: 0 0 auto; width: 76px; height: 50px; border: 1px solid #e1dfdd; border-radius: 2px;
  background: #fff; cursor: pointer; display: flex; flex-direction: column;
  align-items: center; justify-content: center; gap: 3px; padding: 2px;
}
.style-card:hover { border-color: #2b579a; }
.style-card.active { border-color: #2b579a; box-shadow: inset 0 0 0 1px #2b579a; background: #eef3fb; }
.style-card .preview { font-size: 13px; line-height: 1; color: #323130; }
.style-card .name { font-size: 10px; color: #605e5c; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 72px; }

/* work area: optional outline drawer + ruler + the scrolling page canvas */
#workarea { flex: 1 1 auto; min-height: 0; display: flex; }
#editorpane { flex: 1 1 auto; min-width: 0; min-height: 0; display: flex; flex-direction: column; }
/* touch-action: keep one/two-finger PAN (scroll) but disable the browser's
   native pinch-zoom and double-tap-zoom on the document, so our own
   pinch-to-zoom handler (index.ts) owns those gestures. The toolbar/status bar
   keep default touch-action so the ribbon scrolls and taps normally. */
#app { flex: 1 1 auto; min-height: 0; min-width: 0; overflow: auto; background: #e8eaed; position: relative; touch-action: pan-x pan-y; }

/* horizontal ruler (inch ticks, margin shading, draggable indent markers) */
#ruler { flex: 0 0 22px; height: 22px; position: relative; background: #e8eaed; border-bottom: 1px solid #d2d0ce; overflow: hidden; }
#ruler.hidden { display: none; }
#ruler canvas { position: absolute; inset: 0; }
#ruler .ruler-marker { position: absolute; width: 0; height: 0; cursor: ew-resize; z-index: 2; }
/* left-indent: bottom-pointing triangle sitting on the baseline */
#ruler .ruler-left { bottom: 0; border-left: 5px solid transparent; border-right: 5px solid transparent; border-bottom: 7px solid #5b6b8c; }
/* first-line-indent: top-pointing triangle */
#ruler .ruler-first { top: 0; border-left: 5px solid transparent; border-right: 5px solid transparent; border-top: 7px solid #5b6b8c; }

/* ===== Outline / Navigation drawer ================================== */
#outline {
  flex: 0 0 264px; width: 264px; min-height: 0; overflow-y: auto;
  background: #fff; border-right: 1px solid #e1dfdd; display: none;
}
#outline.open { display: block; }
#outline .outline-head {
  position: sticky; top: 0; background: #fff; z-index: 1;
  display: flex; align-items: center; justify-content: space-between;
  padding: 9px 8px 9px 14px; border-bottom: 1px solid #e1dfdd;
  font-size: 13px; font-weight: 600; color: #323130;
}
#outline .outline-head button {
  border: none; background: transparent; cursor: pointer; color: #605e5c;
  width: 24px; height: 24px; border-radius: 4px; font-size: 17px; line-height: 1;
}
#outline .outline-head button:hover { background: #f3f2f1; }
#outline-list { padding: 4px 0 12px; }
.outline-item {
  display: block; width: 100%; box-sizing: border-box; text-align: left;
  border: none; border-left: 3px solid transparent; background: transparent; cursor: pointer;
  padding: 4px 12px; color: #323130; font-size: 13px; line-height: 1.35;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}
.outline-item:hover { background: #f3f2f1; }
.outline-item.active { background: #eef3fb; border-left-color: #2b579a; color: #2b579a; font-weight: 600; }
.outline-empty { padding: 14px; color: #80868b; font-size: 12px; line-height: 1.4; }

/* collapsed ribbon: keep the tab strip, hide the body */
#toolbar.collapsed .rib-bodies { display: none; }

/* ===== Status bar ================================================== */
#statusbar {
  flex: 0 0 auto; height: 24px; display: flex; align-items: center; justify-content: space-between;
  background: #2b579a; color: #fff; font-size: 12px; padding: 0 10px; user-select: none;
}
#statusbar .sb-left, #statusbar .sb-right { display: flex; align-items: center; gap: 16px; }
#statusbar .sb-right { gap: 8px; }
#statusbar .sb-item { white-space: nowrap; }
#statusbar .sb-sep { width: 1px; height: 14px; background: rgba(255,255,255,0.35); }
#statusbar button.sb-btn {
  background: transparent; border: none; color: #fff; cursor: pointer; font-size: 15px;
  width: 22px; height: 20px; border-radius: 3px; line-height: 1; padding: 0;
}
#statusbar button.sb-btn:hover { background: rgba(255,255,255,0.18); }
#statusbar input[type="range"] { width: 120px; cursor: pointer; accent-color: #fff; }
#statusbar .sb-zoom { min-width: 38px; text-align: right; }

/* ===== Popovers (palettes, menus, pickers, dialogs) ================ */
.cw-pop {
  position: fixed; background: #fff; border: 1px solid #c8c6c4; border-radius: 6px;
  box-shadow: 0 4px 18px rgba(0,0,0,0.18); z-index: 50; font-size: 13px; padding: 4px;
}
.cw-menu { display: flex; flex-direction: column; min-width: 168px; }
.cw-menu button {
  display: flex; align-items: center; gap: 10px; width: 100%; text-align: left;
  border: none; background: transparent; cursor: pointer; padding: 7px 10px; border-radius: 4px;
  font: inherit; color: #323130;
}
.cw-menu button:hover { background: #f3f2f1; }
.cw-menu .check { width: 14px; color: #2b579a; }
.cw-menu .sample { color: #605e5c; }
.cw-pop-title { font-size: 11px; color: #605e5c; padding: 4px 6px 2px; }
.cw-swatches { display: grid; grid-template-columns: repeat(6, 22px); gap: 4px; padding: 4px 6px; }
.cw-swatches button { width: 22px; height: 22px; border: 1px solid rgba(0,0,0,0.18); border-radius: 3px; cursor: pointer; padding: 0; }
.cw-swatches button:hover { outline: 2px solid #2b579a; outline-offset: 1px; }
.cw-pop .pop-action {
  display: flex; align-items: center; gap: 8px; width: 100%; text-align: left; border: none;
  background: transparent; cursor: pointer; padding: 7px 8px; border-radius: 4px; font: inherit; color: #323130;
  border-top: 1px solid #edebe9; margin-top: 4px;
}
.cw-pop .pop-action:hover { background: #f3f2f1; }
.cw-grid { display: grid; gap: 2px; padding: 8px 8px 2px; }
.cw-grid .cell { width: 15px; height: 15px; border: 1px solid #c8c6c4; background: #fff; }
.cw-grid .cell.on { background: #cfe3fb; border-color: #2b579a; }
.cw-grid-label { text-align: center; font-size: 12px; color: #605e5c; padding: 4px 0 6px; }
.cw-dialog { display: flex; flex-direction: column; gap: 8px; padding: 10px; min-width: 248px; }
.cw-dialog label { font-size: 12px; color: #605e5c; display: flex; flex-direction: column; gap: 3px; }
.cw-dialog input { height: 28px; border: 1px solid #c8c6c4; border-radius: 4px; padding: 0 8px; font: inherit; font-size: 13px; }
.cw-dialog .row { display: flex; justify-content: flex-end; gap: 6px; margin-top: 2px; }
.cw-dialog button { height: 28px; border: 1px solid #c8c6c4; background: #fff; border-radius: 4px; cursor: pointer; padding: 0 12px; font: inherit; font-size: 13px; }
.cw-dialog button.primary { background: #2b579a; color: #fff; border-color: #2b579a; }
.cw-dialog button.danger { color: #a4262c; }

/* ===== Busy overlay (docx import / join / publish) ================= */
.cw-loading-overlay {
  position: fixed; inset: 0; z-index: 70; background: rgba(0,0,0,0.25);
  display: flex; align-items: center; justify-content: center;
}
.cw-loading-card {
  display: flex; flex-direction: column; align-items: center; gap: 12px;
  background: #fff; border: 1px solid #c8c6c4; border-radius: 8px;
  box-shadow: 0 8px 30px rgba(0,0,0,0.25); padding: 22px 28px; min-width: 200px;
}
.cw-spinner {
  width: 30px; height: 30px; border-radius: 50%;
  border: 3px solid #e1dfdd; border-top-color: #2b579a;
  animation: cw-spin 0.8s linear infinite;
}
@keyframes cw-spin { to { transform: rotate(360deg); } }
.cw-loading-label { font-size: 13px; color: #323130; text-align: center; }
.cw-progress { width: 180px; height: 4px; background: #e1dfdd; border-radius: 2px; overflow: hidden; }
.cw-progress-bar { height: 100%; width: 0%; background: #2b579a; transition: width 0.15s ease; }

/* floating mini-toolbar shown above a selected image */
#img-toolbar {
  position: fixed; display: none; align-items: center; gap: 2px; z-index: 40;
  background: #fff; border: 1px solid #c8c6c4; border-radius: 6px; padding: 3px;
  box-shadow: 0 3px 12px rgba(0,0,0,0.18);
}
#img-toolbar button {
  width: 28px; height: 28px; display: inline-flex; align-items: center; justify-content: center;
  border: 1px solid transparent; border-radius: 4px; background: transparent; cursor: pointer; color: #323130;
}
#img-toolbar button:hover { background: #e1dfdd; }
#img-toolbar button.danger:hover { background: #fde7e9; color: #a4262c; }
#img-toolbar button svg { width: 16px; height: 16px; }
#img-toolbar .sep { width: 1px; height: 18px; background: #e1dfdd; margin: 0 2px; }

/* ===== Mobile / touch responsive layer ============================== */
/* Activates on touch devices (coarse primary pointer) OR narrow screens.
   Desktop is untouched outside this block. Strategy: collapse the ribbon to one
   horizontally-scrollable row, hide group captions, grow touch targets to ~40px,
   turn the 264px outline into an overlay drawer, and clamp floating panels to
   the viewport (a bottom sheet) so they never render off-screen. */
@media (pointer: coarse), (max-width: 760px) {
  /* Ribbon body: single scrollable row instead of a tall multi-group block. */
  .rib-panel { min-height: 0; overflow-x: auto; overflow-y: hidden; flex-wrap: nowrap; -webkit-overflow-scrolling: touch; }
  .rib-label { display: none; }
  .rib-group { padding: 4px 6px; }
  /* Touch targets (Apple/Material guideline is ~40-48px). */
  #toolbar button.rib-btn { min-width: 40px; height: 40px; }
  #toolbar button.rib-btn svg { width: 18px; height: 18px; }
  .rib-tab { padding: 9px 14px; font-size: 14px; }
  #toolbar select, #toolbar input[type="number"] { height: 36px; font-size: 14px; }

  /* Outline: overlay the page instead of stealing 264px of a ~360px screen. */
  #workarea { position: relative; }
  #outline { position: absolute; left: 0; top: 0; bottom: 0; height: auto; width: min(264px, 80vw); z-index: 30; box-shadow: 2px 0 16px rgba(0,0,0,0.18); }

  /* Status bar: tappable zoom controls. */
  #statusbar { height: 36px; }
  #statusbar button.sb-btn { width: 32px; height: 30px; font-size: 17px; }
  #statusbar input[type="range"] { width: 96px; }

  /* Floating panels (Page Setup, Find) → bottom sheet; Activity → full-width.
     !important overrides the inline position/size set in editorApp.ts. */
  .cw-float-panel { left: 8px !important; right: 8px !important; top: auto !important; bottom: 8px !important; width: auto !important; max-width: none !important; max-height: 60vh; overflow: auto; }
  .cw-float-drawer { width: 100% !important; }
  #img-toolbar button { width: 34px; height: 34px; }

  /* Image resize handles: 8px dots are unhittable with a finger — an invisible
     ::before pads the touch target to ~24px without changing the visual size. */
  .cw-obj-handle::before { content: ""; position: absolute; inset: -8px; }
}
`;

/** Inject the component stylesheet once per document (idempotent). */
export function ensureWordCanvasStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.appendChild(style);
}
