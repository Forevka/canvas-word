// Component CSS for the editor chrome (ribbon, ruler, outline, status bar,
// popovers). Injected once by WordCanvas so the package is self-contained — the
// host page needs no stylesheet. The only change from the original index.html
// <style> is that the page-level `html, body` rules become a `.wordcanvas-root`
// container rule, so embedding doesn't restyle the host's body.
//
// Structural selectors are CLASS-based (.cw-toolbar/.cw-app/.cw-ruler/
// .cw-outline/.cw-statusbar), scoped by the per-instance .wordcanvas-root, so
// multiple editors can coexist on one page without id collisions. The stylesheet
// itself is shared (injected once, keyed by STYLE_ID) — it styles every instance.

const STYLE_ID = "wordcanvas-styles";

const CSS = `
.wordcanvas-root {
  display: flex; flex-direction: column; height: 100%; min-height: 0; overflow-x: hidden;
  font-family: "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
  color: #323130;
}

/* ===== Word-style ribbon ============================================ */
.cw-toolbar {
  flex: 0 0 auto; display: flex; flex-direction: column;
  background: #f3f2f1; border-bottom: 1px solid #e1dfdd; user-select: none;
}

/* --- tab strip --- */
.rib-tabs {
  display: flex; align-items: flex-end; gap: 1px; height: 32px;
  padding: 4px 8px 0; background: #f3f2f1;
}
/* Tab buttons scroll horizontally on overflow; the right-side cluster
   (.cw-header-review + collapse chevron) stays pinned as siblings of .rib-tabs. */
.rib-tab-scroll {
  flex: 1 1 auto; min-width: 0; display: flex; align-items: flex-end; gap: 1px;
  overflow-x: auto; overflow-y: hidden; flex-wrap: nowrap; scrollbar-width: none;
}
.rib-tab-scroll::-webkit-scrollbar { display: none; }
.rib-tab {
  flex: 0 0 auto; /* never shrink — tabs scroll instead of squashing */
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
/* contextual tab (Table / Picture / Shape Tools): appears only on selection —
   tinted so it reads as a temporary, selection-scoped tab, not a permanent one */
.rib-tab.rib-tab-ctx { color: #8250df; }
.rib-tab.rib-tab-ctx.active { color: #8250df; }
.rib-tab.rib-tab-ctx.active::after { background: #8250df; }

/* tab-strip overflow "⋯" — appears when the tabs don't fit (critique R3) */
.rib-tab-overflow {
  flex: 0 0 auto; align-self: flex-end; border: none; background: transparent; cursor: pointer;
  color: #605e5c; font-size: 16px; line-height: 1; padding: 3px 8px 7px; border-radius: 4px;
}
.rib-tab-overflow:hover { background: #eceae9; }

/* review controls docked in the ribbon header (right of the tab strip) */
.cw-header-review { margin-left: auto; display: flex; align-items: center; gap: 8px; padding-bottom: 3px; }
.cw-mode-select {
  height: 26px; border: 1px solid #d2d0ce; border-radius: 5px; background: #fff;
  font: inherit; font-size: 12.5px; color: #323130; padding: 0 6px; cursor: pointer;
}
.cw-mode-select:hover { border-color: #b3b0ad; }
.cw-header-btn {
  height: 26px; border: 1px solid #d2d0ce; border-radius: 5px; background: #fff; cursor: pointer;
  font: inherit; font-size: 12.5px; font-weight: 600; color: #2b579a; padding: 0 12px;
  display: inline-flex; align-items: center; gap: 5px;
}
.cw-header-btn:hover { background: #f3f2f1; }
.cw-header-btn.active { background: #2b579a; color: #fff; border-color: #2b579a; }

/* document identity + live save state, pinned to the LEFT of the tab strip so
   "what is this file / is my work safe?" is answerable without opening a menu */
.cw-doc-ident {
  flex: 0 0 auto; align-self: flex-end; display: flex; align-items: center;
  gap: 8px; min-width: 0; max-width: 480px; padding: 0 12px 4px 4px;
}
.cw-doc-title {
  font-size: 13px; font-weight: 600; color: #201f1e; min-width: 0; max-width: 210px;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}
.cw-save { flex: 0 0 auto; display: inline-flex; align-items: center; gap: 5px; cursor: default; }
.cw-save-dot { flex: 0 0 auto; width: 8px; height: 8px; border-radius: 50%; background: #a19f9d; }
.cw-save-dot.dirty { background: #c19c00; }   /* amber — unsaved edits */
.cw-save-dot.saving { background: #2b579a; }  /* blue — in progress */
.cw-save-dot.saved { background: #2e7d32; }   /* green — saved */
.cw-save-dot.clean { background: #a19f9d; }   /* grey — neutral / no target */
.cw-save-text { font-size: 12px; color: #605e5c; white-space: nowrap; }

/* import fidelity badge (Move 3) — the moat made visible: a passive trust signal
   sitting with save state. Faithful (green) when nothing was adapted on import;
   notes (amber) opens a plain-language list of what was preserved-but-adapted. */
.cw-fidelity {
  flex: 0 0 auto; display: inline-flex; align-items: center; gap: 4px;
  border: 1px solid transparent; border-radius: 11px; padding: 1px 9px;
  font-size: 11px; font-weight: 600; line-height: 1.6; cursor: pointer;
  white-space: nowrap; background: none; font-family: inherit;
}
.cw-fidelity.faithful { color: #1b6e2e; background: #e6f4ea; border-color: #b7e0c1; }
.cw-fidelity.notes { color: #8a5300; background: #fdf1dc; border-color: #f0d199; }
.cw-fidelity:hover { filter: brightness(0.97); }
.cw-fidelity-panel {
  position: fixed; z-index: 1200; width: 340px; max-width: 92vw;
  background: #fff; border: 1px solid #e0e0e0; border-radius: 10px;
  box-shadow: 0 12px 40px rgba(0,0,0,.22); padding: 14px 16px;
  font-size: 13px; line-height: 1.5; color: #201f1e;
}
.cw-fidelity-head { font-size: 13px; font-weight: 700; margin-bottom: 4px; }
.cw-fidelity-sub { font-size: 12px; color: #605e5c; margin-bottom: 8px; }
.cw-fidelity-list { margin: 0; padding-left: 18px; display: flex; flex-direction: column; gap: 5px; max-height: 320px; overflow-y: auto; }
.cw-fidelity-list li { font-size: 12px; color: #3c4043; }

/* quick-access toolbar (undo / redo / save) — pinned between the title and the
   tab strip so the most-used commands are one click from any tab */
.cw-qat {
  flex: 0 0 auto; align-self: flex-end; display: flex; align-items: center; gap: 1px;
  padding: 0 4px 3px; margin-left: 2px; border-left: 1px solid #d8d6d4;
}
.cw-qat-btn {
  width: 28px; height: 26px; display: inline-flex; align-items: center; justify-content: center;
  border: none; background: transparent; border-radius: 5px; cursor: pointer; color: #323130; padding: 0;
}
.cw-qat-btn:hover:not(:disabled) { background: #eceae9; }
.cw-qat-btn:disabled { opacity: 0.38; cursor: default; }
.cw-qat-btn svg { width: 16px; height: 16px; }

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
.cw-toolbar button.rib-btn {
  min-width: 26px; height: 26px; display: inline-flex; align-items: center; justify-content: center;
  border: 1px solid transparent; border-radius: 4px; background: transparent; cursor: pointer;
  color: #323130; font-size: 13px; padding: 0 4px; gap: 1px;
}
.cw-toolbar button.rib-btn:hover { background: #e1dfdd; }
.cw-toolbar button.rib-btn:active { background: #d2d0ce; }
.cw-toolbar button.rib-btn.active { background: #cfe3fb; color: #0b57d0; border-color: #b3d3f5; }
.cw-toolbar button.rib-btn.active:hover { background: #bcd8fa; }
.cw-toolbar button.rib-btn.active svg { color: #0b57d0; }
.cw-toolbar button.rib-btn svg { width: 16px; height: 16px; display: block; }
.cw-toolbar button.rib-btn .caret { width: 8px; height: 8px; }

/* large button (Paste): icon over caption */
.cw-toolbar button.rib-big {
  flex-direction: column; height: 100%; min-width: 48px; padding: 4px 6px; gap: 2px; justify-content: center;
}
.cw-toolbar button.rib-big svg { width: 26px; height: 26px; }
.cw-toolbar button.rib-big .big-cap { font-size: 11px; display: flex; align-items: center; gap: 1px; }

/* swatch button: icon row + colour underline */
.cw-toolbar button.rib-swatch { flex-direction: column; gap: 0; padding: 2px 3px; }
.cw-toolbar button.rib-swatch .row { display: flex; align-items: center; gap: 1px; }
.cw-toolbar button.rib-swatch .bar { width: 18px; height: 3px; margin-top: 1px; border-radius: 1px; }

/* disabled stub (feature not supported by the engine yet) */
.cw-toolbar button.rib-btn:disabled { opacity: 0.38; cursor: default; }
.cw-toolbar button.rib-btn:disabled:hover { background: transparent; }

/* form controls */
.cw-toolbar select, .cw-toolbar input[type="number"] {
  height: 24px; border: 1px solid #c8c6c4; border-radius: 3px; background: #fff;
  font: inherit; font-size: 13px; color: #323130;
}
.cw-toolbar select:hover, .cw-toolbar input[type="number"]:hover { border-color: #8a8886; }

/* styles gallery. Width is pinned to a WHOLE number of cards so the horizontal
   scroll never rests showing a clipped half-card (critique C3): with a 76px card
   + 4px gap and box-sizing: border-box, the 320px inner width holds exactly four
   cards (4·76 + 3·4 = 316) with the fifth starting past the edge. scroll-snap
   keeps scrolled positions on card boundaries too. */
.rib-gallery {
  box-sizing: border-box;
  display: flex; align-items: center; gap: 4px; height: 64px; padding: 0 4px;
  border: 1px solid #c8c6c4; border-radius: 4px; background: #fff;
  overflow-x: auto; max-width: 320px; scroll-snap-type: x proximity;
}
.rib-gallery::-webkit-scrollbar { height: 8px; }
.rib-gallery::-webkit-scrollbar-thumb { background: #c1c1c1; border-radius: 4px; }
.style-card {
  flex: 0 0 auto; width: 76px; height: 50px; border: 1px solid #e1dfdd; border-radius: 2px;
  background: #fff; cursor: pointer; display: flex; flex-direction: column;
  align-items: center; justify-content: center; gap: 3px; padding: 2px; scroll-snap-align: start;
}
.style-card:hover { border-color: #2b579a; }
.style-card.active { border-color: #2b579a; box-shadow: inset 0 0 0 1px #2b579a; background: #eef3fb; }
/* position: relative + a fixed box so the canvas sample (rendered async by the
   child document) can overlay a plain-text fallback — no flash of blank rows
   before the canvas paints (critique V5). */
.style-card .preview {
  position: relative; width: 70px; height: 24px; overflow: hidden;
  display: flex; align-items: center; justify-content: center;
  font-size: 13px; line-height: 1; color: #323130;
}
/* the child-document canvas host sets an inline position:relative, so override
   it (!important beats the inline style) to overlay the text fallback. */
.style-card .preview > div { position: absolute !important; inset: 0; }
.style-card .name { font-size: 10px; color: #605e5c; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 72px; }

/* work area: optional outline drawer + ruler + the scrolling page canvas.
   position: relative anchors the right-docked drawers (.cw-float-drawer) to this
   row, so they sit below the ribbon and above the status bar rather than over
   the window / an embedding host page. */
.cw-workarea { flex: 1 1 auto; min-height: 0; display: flex; position: relative; }
.cw-editorpane { flex: 1 1 auto; min-width: 0; min-height: 0; display: flex; flex-direction: column; }
/* touch-action: keep one/two-finger PAN (scroll) but disable the browser's
   native pinch-zoom and double-tap-zoom on the document, so our own
   pinch-to-zoom handler (index.ts) owns those gestures. The toolbar/status bar
   keep default touch-action so the ribbon scrolls and taps normally. */
.cw-app { flex: 1 1 auto; min-height: 0; min-width: 0; overflow: auto; background: #e8eaed; position: relative; touch-action: pan-x pan-y; }

/* ruler row: top-left corner spacer (over the vertical ruler) + horizontal ruler */
.cw-ruler-row { flex: 0 0 22px; display: flex; }
.cw-ruler-row.hidden { display: none; }
.cw-ruler-corner { flex: 0 0 22px; width: 22px; background: #e8eaed; border-bottom: 1px solid #d2d0ce; border-right: 1px solid #d2d0ce; }
.cw-ruler-corner.hidden { display: none; }
.cw-main-row { flex: 1 1 auto; min-height: 0; min-width: 0; display: flex; }
/* horizontal ruler (inch ticks, margin shading, draggable indent markers) */
.cw-ruler { flex: 1 1 auto; height: 22px; position: relative; background: #e8eaed; border-bottom: 1px solid #d2d0ce; overflow: hidden; }
.cw-ruler.hidden { display: none; }
.cw-ruler canvas { position: absolute; inset: 0; }
/* vertical ruler (inch ticks + top/bottom margin shading down the left edge) */
.cw-vruler { flex: 0 0 22px; width: 22px; position: relative; background: #e8eaed; border-right: 1px solid #d2d0ce; overflow: hidden; }
.cw-vruler.hidden { display: none; }
.cw-vruler canvas { position: absolute; inset: 0; }
/* margin handles: right-pointing triangles at the top/bottom content boundaries */
.cw-vruler .vruler-marker { position: absolute; left: 0; width: 0; height: 0; cursor: ns-resize; z-index: 2; border-top: 5px solid transparent; border-bottom: 5px solid transparent; border-left: 7px solid #5b6b8c; }
.cw-ruler .ruler-marker { position: absolute; width: 0; height: 0; cursor: ew-resize; z-index: 2; }
/* left-indent: bottom-pointing triangle sitting on the baseline */
.cw-ruler .ruler-left { bottom: 0; border-left: 5px solid transparent; border-right: 5px solid transparent; border-bottom: 7px solid #5b6b8c; }
/* first-line-indent: top-pointing triangle */
.cw-ruler .ruler-first { top: 0; border-left: 5px solid transparent; border-right: 5px solid transparent; border-top: 7px solid #5b6b8c; }

/* ===== Outline / Navigation drawer ================================== */
/* Column: head + filter are fixed; only the list scrolls, so the drag-resize
   handle (position:absolute) stays pinned to the right edge (O5). */
.cw-outline {
  position: relative; flex: 0 0 264px; width: 264px; min-height: 0;
  background: #fff; border-right: 1px solid #e1dfdd; display: none;
  flex-direction: column; overflow: hidden;
}
.cw-outline.open { display: flex; }
.cw-outline .outline-head {
  flex: 0 0 auto; background: #fff;
  display: flex; align-items: center; justify-content: space-between;
  padding: 9px 8px 9px 14px; border-bottom: 1px solid #e1dfdd;
  font-size: 13px; font-weight: 600; color: #323130;
}
.cw-outline .outline-head button {
  border: none; background: transparent; cursor: pointer; color: #605e5c;
  width: 24px; height: 24px; border-radius: 4px; font-size: 17px; line-height: 1;
}
.cw-outline .outline-head button:hover { background: #f3f2f1; }
.cw-outline-filter { flex: 0 0 auto; padding: 6px 10px; border-bottom: 1px solid #f0eeee; }
.cw-outline-filter input {
  width: 100%; box-sizing: border-box; height: 26px; border: 1px solid #d2d0ce;
  border-radius: 5px; padding: 0 8px; font: inherit; font-size: 12.5px; color: #323130;
}
.cw-outline-filter input:focus { outline: none; border-color: #2b579a; box-shadow: 0 0 0 1px #2b579a; }
.cw-outline-list { flex: 1 1 auto; overflow-y: auto; padding: 4px 0 12px; }
.cw-outline-resize {
  position: absolute; top: 0; right: 0; width: 6px; height: 100%; cursor: col-resize; z-index: 3;
}
.cw-outline-resize:hover { background: linear-gradient(90deg, transparent, rgba(43,87,154,0.25)); }
.outline-item {
  display: flex; align-items: center; gap: 4px; width: 100%; box-sizing: border-box; text-align: left;
  border: none; border-left: 3px solid transparent; background: transparent; cursor: pointer;
  padding: 3px 10px 3px 0; color: #323130; line-height: 1.3;
}
.outline-item:hover { background: #f3f2f1; }
.outline-item.active { background: #eef3fb; border-left-color: #2b579a; }
.outline-item.active .outline-label { color: #2b579a; }
.outline-chevron {
  flex: 0 0 auto; width: 16px; height: 16px; border: none; background: transparent; cursor: pointer;
  color: #80868b; font-size: 9px; line-height: 1; padding: 0; border-radius: 3px;
}
.outline-chevron:hover { background: #e6e4e2; color: #323130; }
.outline-chevron.leaf { visibility: hidden; }
.outline-label { flex: 1 1 auto; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.outline-page { flex: 0 0 auto; color: #a19f9d; font-size: 11px; font-variant-numeric: tabular-nums; padding-left: 6px; }
/* Level styling (O4): weight/size/colour differentiate H1..H3+, not just indent. */
.outline-item.lvl-0 .outline-label, .outline-item.lvl-1 .outline-label { font-weight: 600; font-size: 13px; color: #201f1e; }
.outline-item.lvl-2 .outline-label { font-weight: 500; font-size: 12.5px; color: #3c4043; }
.outline-item.lvl-3 .outline-label { font-weight: 400; font-size: 12px; color: #605e5c; }
.outline-empty { padding: 14px; color: #80868b; font-size: 12px; line-height: 1.4; }
/* drag-to-reorder (O1): dragged row dims; a blue rule marks the drop position */
.outline-item.cw-outline-dragging { opacity: 0.4; }
.cw-outline-drop { position: absolute; left: 6px; right: 6px; height: 2px; background: #2b579a; z-index: 4; pointer-events: none; }
:root[data-theme="dark"] .cw-outline-drop { background: #7cb0ff; }

/* ===== Navigator: 48px icon rail + swappable tab panels (row 17 / S1) === */
.cw-outline.cw-navigator { flex-direction: row; }
.cw-nav-rail {
  flex: 0 0 48px; display: flex; flex-direction: column; align-items: center; gap: 2px;
  padding: 6px 0; background: #f3f2f1; border-right: 1px solid #e1dfdd;
}
.cw-nav-tab {
  width: 36px; height: 36px; display: inline-flex; align-items: center; justify-content: center;
  border: none; background: transparent; border-radius: 6px; cursor: pointer; color: #605e5c; padding: 0;
}
.cw-nav-tab svg { width: 18px; height: 18px; }
.cw-nav-tab:hover { background: #e6e4e2; color: #323130; }
.cw-nav-tab.active { background: #fff; color: #2b579a; box-shadow: inset 2px 0 0 #2b579a; }
.cw-nav-content { flex: 1 1 auto; min-width: 0; display: flex; flex-direction: column; }
.cw-nav-panels { flex: 1 1 auto; min-height: 0; position: relative; }
.cw-nav-panel { display: none; position: absolute; inset: 0; flex-direction: column; min-height: 0; }
.cw-nav-panel.active { display: flex; }
.cw-nav-action {
  flex: 0 0 auto; margin: 8px 10px 4px; height: 28px; border: 1px solid #d2d0ce; border-radius: 6px;
  background: #fff; cursor: pointer; font: inherit; font-size: 12.5px; color: #2b579a; font-weight: 600;
}
.cw-nav-action:hover { background: #f3f2f1; }
:root[data-theme="dark"] .cw-nav-rail { background: #26282c; border-color: #34373c; }
:root[data-theme="dark"] .cw-nav-tab { color: #9aa0a6; }
:root[data-theme="dark"] .cw-nav-tab:hover { background: #34373c; color: #e6e6e6; }
:root[data-theme="dark"] .cw-nav-tab.active { background: #1e1f22; color: #7cb0ff; box-shadow: inset 2px 0 0 #7cb0ff; }
:root[data-theme="dark"] .cw-nav-action { background: #2a2c30; color: #7cb0ff; border-color: #4a4a4a; }

/* Styles tab (S2): live style cards — hover-preview, click-apply, right-click */
.cw-styles-list { flex: 1 1 auto; min-height: 0; overflow-y: auto; padding: 4px 6px 8px; display: flex; flex-direction: column; gap: 2px; }
.cw-style-card {
  display: flex; align-items: center; gap: 8px; width: 100%; text-align: left;
  border: 1px solid transparent; border-radius: 6px; background: none; cursor: pointer;
  padding: 4px 8px; font: inherit; color: #201f1e;
}
.cw-style-card:hover { background: #eef3fb; border-color: #cfe0fa; }
.cw-style-card.current { background: #eef3fb; border-color: #2b579a; }
.cw-style-swatch { flex: 0 0 auto; width: 88px; height: 24px; overflow: hidden; border-radius: 3px; }
.cw-style-name { flex: 1 1 auto; min-width: 0; font-size: 12.5px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.cw-style-count { flex: 0 0 auto; font-size: 11px; color: #80868b; min-width: 14px; text-align: right; }
.cw-styles-chip {
  flex: 0 0 auto; margin: 8px 8px 4px; display: flex; align-items: center; gap: 8px;
  padding: 5px 8px; border-radius: 6px; background: #fdf1dc; border: 1px solid #f0d199; font-size: 11.5px; color: #8a5300;
}
.cw-styles-chip span { flex: 1 1 auto; }
.cw-styles-chip-clear {
  flex: 0 0 auto; border: 1px solid #e0b866; border-radius: 5px; background: #fff; cursor: pointer;
  font: inherit; font-size: 11px; padding: 2px 8px; color: #8a5300;
}
.cw-styles-chip-clear:hover { background: #fbe9c8; }
:root[data-theme="dark"] .cw-style-card { color: #e6e6e6; }
:root[data-theme="dark"] .cw-style-card:hover { background: #2c3340; border-color: #3a4a63; }
:root[data-theme="dark"] .cw-style-card.current { background: #2c3340; border-color: #7cb0ff; }
:root[data-theme="dark"] .cw-style-count { color: #9aa0a6; }
:root[data-theme="dark"] .cw-styles-chip { background: #2e2517; border-color: #5a4a2a; color: #e6bd73; }
:root[data-theme="dark"] .cw-styles-chip-clear { background: #26282c; border-color: #5a4a2a; color: #e6bd73; }

/* ===== Review pane (track changes + comments) ======================= */
/* ===== Inspector: right-docked, selection-aware property sheet (Move 2) ===== */
.cw-inspector {
  flex: 0 0 300px; width: 300px; min-height: 0; display: none;
  flex-direction: column; background: #f7f8fa; border-left: 1px solid #e1dfdd;
  font-size: 13px; color: #202124;
}
.cw-inspector.open { display: flex; }
.cw-insp-head { display: flex; align-items: center; gap: 8px; padding: 10px 10px 10px 14px; background: #fff; border-bottom: 1px solid #e8eaed; }
.cw-insp-title { font-size: 14px; font-weight: 600; flex: 1 1 auto; }
.cw-insp-close { border: none; background: transparent; cursor: pointer; color: #5f6368; width: 28px; height: 28px; border-radius: 50%; font-size: 18px; line-height: 1; }
.cw-insp-close:hover { background: #f1f3f4; }
.cw-insp-body { flex: 1 1 auto; min-height: 0; overflow-y: auto; padding: 4px 0 12px; }
.cw-insp-section { border-bottom: 1px solid #e8eaed; padding: 10px 14px 12px; }
.cw-insp-section h3 { margin: 0 0 8px; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .05em; color: #80868b; }
.cw-insp-row { display: flex; align-items: center; gap: 8px; margin-bottom: 8px; }
.cw-insp-row:last-child { margin-bottom: 0; }
.cw-insp-label { flex: 0 0 74px; font-size: 12px; color: #5f6368; }
.cw-insp-row > select, .cw-insp-row > input[type="number"], .cw-insp-row > input[type="text"] { flex: 1 1 auto; min-width: 0; height: 28px; border: 1px solid #d0d4d9; border-radius: 6px; padding: 0 8px; font: inherit; font-size: 12.5px; }
.cw-insp-btns { display: inline-flex; gap: 2px; flex: 1 1 auto; flex-wrap: wrap; }
.cw-insp-pair { display: flex; gap: 6px; flex: 1 1 auto; min-width: 0; }
.cw-insp-pair > input { flex: 1 1 0; min-width: 0; height: 28px; border: 1px solid #d0d4d9; border-radius: 6px; padding: 0 8px; font: inherit; font-size: 12.5px; }
.cw-insp-tgl { min-width: 30px; height: 28px; border: 1px solid #d0d4d9; background: #fff; cursor: pointer; font: inherit; font-size: 12.5px; border-radius: 6px; color: #3c4043; padding: 0 6px; }
.cw-insp-tgl.on { background: #eef3fb; border-color: #2b579a; color: #2b579a; }
.cw-insp-tgl:hover:not(.on) { background: #f3f2f1; }
.cw-insp-color { width: 34px; height: 28px; padding: 0; border: 1px solid #d0d4d9; border-radius: 6px; cursor: pointer; background: #fff; }
.cw-insp-empty { padding: 24px 16px; color: #80868b; font-size: 12.5px; line-height: 1.5; text-align: center; }
:root[data-theme="dark"] .cw-inspector { background: #202124; border-color: #34373c; color: #e6e6e6; }
:root[data-theme="dark"] .cw-insp-head { background: #26282c; border-color: #34373c; }
:root[data-theme="dark"] .cw-insp-section { border-color: #34373c; }
:root[data-theme="dark"] .cw-insp-row > select, :root[data-theme="dark"] .cw-insp-row > input, :root[data-theme="dark"] .cw-insp-pair > input { background: #2a2c30; color: #e6e6e6; border-color: #4a4a4a; }
:root[data-theme="dark"] .cw-insp-tgl { background: #2a2c30; color: #e0e0e0; border-color: #4a4a4a; }
:root[data-theme="dark"] .cw-insp-tgl.on { background: #2c3340; color: #7cb0ff; border-color: #7cb0ff; }
:root[data-theme="dark"] .cw-insp-label, :root[data-theme="dark"] .cw-insp-title { color: #c8ccd0; }
:root[data-theme="dark"] .cw-insp-close { color: #9aa0a6; }
:root[data-theme="dark"] .cw-insp-close:hover { background: #34373c; }

.cw-review {
  flex: 0 0 320px; width: 320px; min-height: 0; display: none;
  flex-direction: column; background: #f7f8fa; border-left: 1px solid #e1dfdd;
  font-size: 13px; color: #202124;
}
.cw-review.open { display: flex; }
.cw-review-head {
  display: flex; align-items: center; gap: 8px;
  padding: 10px 10px 10px 14px; background: #fff; border-bottom: 1px solid #e8eaed;
}
.cw-review-head .cw-review-title { font-size: 14px; font-weight: 600; color: #202124; }
.cw-review-head .cw-review-close {
  margin-left: auto; border: none; background: transparent; cursor: pointer;
  color: #5f6368; width: 28px; height: 28px; border-radius: 50%; font-size: 18px; line-height: 1;
}
.cw-review-head .cw-review-close:hover { background: #f1f3f4; }
/* tab strip */
.cw-review-tabs { display: flex; background: #fff; border-bottom: 1px solid #e8eaed; padding: 0 8px; }
.cw-review-tab {
  flex: 1 1 0; border: none; background: transparent; cursor: pointer;
  padding: 9px 4px 8px; font-size: 13px; font-weight: 600; color: #5f6368;
  border-bottom: 2px solid transparent; display: flex; align-items: center; justify-content: center; gap: 6px;
}
.cw-review-tab:hover { color: #202124; }
.cw-review-tab.active { color: #1a73e8; border-bottom-color: #1a73e8; }
.cw-review-tab .cw-pill {
  background: #e8eaed; color: #5f6368; border-radius: 10px; padding: 0 7px;
  font-size: 11px; font-weight: 700; min-width: 18px; text-align: center;
}
.cw-review-tab.active .cw-pill { background: #d2e3fc; color: #1a73e8; }
/* sticky bulk-action bar */
.cw-review-actions {
  display: flex; gap: 8px; padding: 8px 12px; background: #f7f8fa;
  border-bottom: 1px solid #e8eaed; position: sticky; top: 0; z-index: 1;
}
.cw-review-body { flex: 1 1 auto; overflow-y: auto; padding: 10px 12px; }
.cw-review-empty { padding: 28px 16px; text-align: center; color: #80868b; font-size: 12.5px; line-height: 1.5; }
.cw-review-empty .cw-review-empty-ico { font-size: 26px; display: block; margin-bottom: 8px; opacity: .6; }

/* ===== Right-docked drawers (Bookmarks / Activity) ================== */
/* Absolutely positioned inside .cw-workarea (which is position: relative), so a
   drawer overlays the editor pane but stays BELOW the ribbon and ABOVE the
   status bar — and, in an embed, never paints over the host page. Visibility is
   toggled via the inline display style (none / flex) in editorApp.ts. Was
   previously a position:fixed window overlay that covered the ribbon and, on
   coarse pointers, the whole viewport. */
.cw-float-drawer {
  position: absolute; top: 0; right: 0; bottom: 0;
  width: min(300px, 92%); z-index: 40;
  background: #fff; border-left: 1px solid #e1dfdd;
  box-shadow: -4px 0 16px rgba(0,0,0,0.08);
  flex-direction: column; font-size: 13px;
}

/* buttons */
.cw-btn {
  border: 1px solid #dadce0; background: #fff; color: #3c4043; cursor: pointer;
  border-radius: 6px; padding: 5px 12px; font-size: 12.5px; font-weight: 600;
  display: inline-flex; align-items: center; gap: 5px; line-height: 1;
}
.cw-btn:hover { background: #f8f9fa; box-shadow: 0 1px 2px rgba(60,64,67,.15); }
.cw-btn.cw-btn-primary { background: #1a73e8; border-color: #1a73e8; color: #fff; }
.cw-btn.cw-btn-primary:hover { background: #1b66c9; }
.cw-btn.cw-btn-accept { color: #137333; border-color: #c6e0c9; background: #e6f4ea; }
.cw-btn.cw-btn-accept:hover { background: #d3ecd9; }
.cw-btn.cw-btn-reject { color: #c5221f; border-color: #f3c7c5; background: #fce8e6; }
.cw-btn.cw-btn-reject:hover { background: #f9d6d3; }
.cw-btn.cw-btn-ghost { border-color: transparent; background: transparent; color: #1a73e8; padding: 4px 6px; }
.cw-btn.cw-btn-ghost:hover { background: #e8f0fe; box-shadow: none; }
.cw-btn-sm { padding: 4px 9px; font-size: 12px; }

/* suggestion card */
.cw-sug {
  background: #fff; border: 1px solid #e8eaed; border-left: 3px solid var(--cw-author, #1a73e8);
  border-radius: 8px; padding: 10px 12px; margin-bottom: 8px; box-shadow: 0 1px 2px rgba(60,64,67,.06);
  cursor: pointer;
}
.cw-sug:hover { background: #f8f9fa; border-color: #d2e3fc; }
.cw-sug-top { display: flex; align-items: center; gap: 8px; margin-bottom: 8px; }
.cw-sug-kind { font-weight: 600; font-size: 12.5px; color: #202124; display: flex; align-items: center; gap: 6px; }
.cw-sug-kind .cw-dot { width: 8px; height: 8px; border-radius: 50%; background: var(--cw-author, #1a73e8); }
.cw-sug-meta { margin-left: auto; color: #80868b; font-size: 11.5px; white-space: nowrap; }
.cw-sug-actions { display: flex; gap: 8px; }

/* avatar + comment card */
.cw-avatar {
  width: 28px; height: 28px; border-radius: 50%; flex: 0 0 28px;
  display: flex; align-items: center; justify-content: center;
  color: #fff; font-size: 12px; font-weight: 700; text-transform: uppercase;
}
.cw-thread {
  background: #fff; border: 1px solid #e8eaed; border-radius: 8px;
  padding: 12px; margin-bottom: 10px; box-shadow: 0 1px 2px rgba(60,64,67,.06);
  cursor: pointer;
}
.cw-thread:hover { border-color: #d2e3fc; }
.cw-thread.resolved { opacity: .62; }
.cw-comment { display: flex; gap: 9px; margin-bottom: 10px; }
.cw-comment:last-of-type { margin-bottom: 6px; }
.cw-comment-main { min-width: 0; flex: 1 1 auto; }
.cw-comment-who { font-weight: 600; font-size: 12.5px; color: #202124; }
.cw-comment-when { color: #80868b; font-size: 11px; margin-left: 6px; font-weight: 400; }
.cw-comment-body { color: #3c4043; font-size: 13px; line-height: 1.45; white-space: pre-wrap; word-wrap: break-word; margin-top: 2px; }
.cw-thread-actions { display: flex; gap: 4px; align-items: center; padding-top: 6px; border-top: 1px solid #f1f3f4; }
.cw-thread-actions .cw-resolved-tag { color: #137333; font-size: 11.5px; font-weight: 600; display: flex; align-items: center; gap: 4px; }

/* inline reply editor inside a thread */
.cw-reply-box { display: none; gap: 8px; margin-top: 8px; }
.cw-reply-box.open { display: flex; }

/* ===== Comment composer bubble (Google-Docs style, floats by selection) === */
.cw-comment-bubble {
  position: absolute; z-index: 60; width: 300px;
  background: #fff; border: 1px solid #e0e0e0; border-radius: 10px;
  box-shadow: 0 6px 22px rgba(60,64,67,.28); padding: 12px;
  display: flex; flex-direction: column; gap: 10px;
}
.cw-comment-bubble .cw-bubble-row { display: flex; gap: 9px; align-items: flex-start; }
.cw-comment-bubble textarea {
  flex: 1 1 auto; resize: none; min-height: 48px; max-height: 180px;
  border: none; outline: none; font: 13px/1.45 inherit; color: #202124;
  padding: 4px 0; background: transparent;
}
.cw-comment-bubble .cw-bubble-actions { display: flex; justify-content: flex-end; gap: 8px; }
.cw-comment-bubble .cw-btn.cw-btn-primary:disabled { opacity: .45; cursor: default; box-shadow: none; }

/* floating "leave a comment" chip shown beside a suggest-mode selection */
.cw-comment-chip {
  position: absolute; z-index: 55; width: 34px; height: 34px; border-radius: 50%;
  border: 1px solid #e0e0e0; background: #fff; cursor: pointer; font-size: 16px;
  box-shadow: 0 2px 8px rgba(60,64,67,.28); display: flex; align-items: center; justify-content: center;
  padding: 0; line-height: 1; transition: transform .08s ease, box-shadow .08s ease;
}
.cw-comment-chip:hover { transform: scale(1.08); box-shadow: 0 3px 12px rgba(60,64,67,.34); background: #f8f9fa; }

/* @-mention autocomplete dropdown + rendered mention chips */
.cw-mention-menu {
  position: fixed; z-index: 70; background: #fff; border: 1px solid #e0e0e0;
  border-radius: 8px; box-shadow: 0 6px 22px rgba(60,64,67,.28); padding: 4px;
  max-height: 240px; overflow-y: auto; font: 13px/1.3 inherit;
}
.cw-mention-item {
  display: flex; align-items: center; gap: 8px; padding: 6px 8px; border-radius: 6px; cursor: pointer; color: #202124;
}
.cw-mention-item:hover, .cw-mention-item.active { background: #e8f0fe; }
.cw-mention-av {
  width: 24px; height: 24px; border-radius: 50%; flex: 0 0 24px;
  display: flex; align-items: center; justify-content: center; color: #fff; font-size: 10px; font-weight: 700; text-transform: uppercase;
}
.cw-mention { color: #1a73e8; background: #e8f0fe; border-radius: 4px; padding: 0 3px; font-weight: 600; }

/* collapsed ribbon: keep the tab strip, hide the body */
.cw-toolbar.collapsed .rib-bodies { display: none; }

/* compact ribbon (container-width driven via ResizeObserver, see editorApp):
   dense single horizontally-scrollable group row with captions hidden. Keyed off
   the editor's own width, so it also engages for narrow embeds on a wide page.
   The pointer:coarse half of the mobile @media block below carries the same body
   rules for touch devices regardless of width. */
.cw-toolbar.compact .rib-panel {
  min-height: 0; overflow-x: auto; overflow-y: hidden;
  flex-wrap: nowrap; -webkit-overflow-scrolling: touch;
}
.cw-toolbar.compact .rib-label { display: none; }
.cw-toolbar.compact .rib-group { padding: 4px 6px; }
.cw-toolbar.compact .rib-panel::-webkit-scrollbar { height: 8px; }
.cw-toolbar.compact .rib-panel::-webkit-scrollbar-thumb { background: #c1c1c1; border-radius: 4px; }

/* ===== Status bar ================================================== */
.cw-statusbar {
  flex: 0 0 auto; height: 24px; display: flex; align-items: center; justify-content: space-between;
  background: #2b579a; color: #fff; font-size: 12px; padding: 0 10px; user-select: none;
}
.cw-statusbar .sb-left, .cw-statusbar .sb-right { display: flex; align-items: center; gap: 16px; }
.cw-statusbar .sb-right { gap: 8px; }
.cw-statusbar .sb-item { white-space: nowrap; }
.cw-statusbar .sb-sep { width: 1px; height: 14px; background: rgba(255,255,255,0.35); }
.cw-statusbar button.sb-btn {
  background: transparent; border: none; color: #fff; cursor: pointer; font-size: 15px;
  width: 22px; height: 20px; border-radius: 3px; line-height: 1; padding: 0;
}
.cw-statusbar button.sb-btn:hover { background: rgba(255,255,255,0.18); }
.cw-statusbar input[type="range"] { width: 120px; cursor: pointer; accent-color: #fff; }
.cw-statusbar .sb-zoom { min-width: 38px; text-align: right; }

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
.cw-pop .pop-row { display: flex; align-items: center; gap: 6px; padding: 4px 6px; }
.cw-pop .pop-row-label { min-width: 52px; font-size: 12px; color: #605e5c; }
.cw-pop .pop-row select {
  flex: 1; height: 24px; font: inherit; font-size: 12px; color: #323130;
  border: 1px solid #c8c6c4; border-radius: 4px; background: #fff; padding: 0 4px; cursor: pointer;
}
.cw-grid { display: grid; gap: 2px; padding: 8px 8px 2px; }
.cw-grid .cell { width: 15px; height: 15px; border: 1px solid #c8c6c4; background: #fff; }
.cw-grid .cell.on { background: #cfe3fb; border-color: #2b579a; }
.cw-grid-label { text-align: center; font-size: 12px; color: #605e5c; padding: 4px 0 6px; }
.cw-grid-cat { font-size: 11px; font-weight: 600; color: #605e5c; padding: 6px 8px 0; }
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

/* The image / shape / table floating object bars now share the cw-ctxbar + cw-iconbar
   base class (see ui/contextToolbar.ts) — no bespoke .cw-img-toolbar block here. */

/* ===== Responsive: narrow-viewport layout + touch hit targets ======== */
/* Two independent concerns, deliberately kept apart. They were once conflated
   under a single @media (pointer: coarse), (max-width: 760px) query whose
   unbounded coarse arm forced the phone layout onto every touchscreen laptop at
   any resolution (opening a drawer then covered the whole viewport). The split:
     • LAYOUT collapse keys on WIDTH only — a touchscreen laptop keeps the
       desktop layout. (The ribbon body ALSO collapses via the container-width
       .compact class driven by the ResizeObserver in editorApp, which covers
       narrow embeds on a wide page.)
     • HIT-TARGET growth keys on any-pointer: coarse — touch capability, not
       screen size — but is bounded to small screens, so a wide mouse-driven
       touchscreen laptop stays compact and the ribbon does not overflow. */

/* Layout: genuinely narrow viewports only. */
@media (max-width: 760px) {
  /* Ribbon body: single scrollable row instead of a tall multi-group block. */
  .rib-panel { min-height: 0; overflow-x: auto; overflow-y: hidden; flex-wrap: nowrap; -webkit-overflow-scrolling: touch; }
  .rib-label { display: none; }
  .rib-group { padding: 4px 6px; }

  /* Keep the title + save dot but drop the save-state words so the pinned
     identity cluster doesn't crowd the tab strip on a phone. */
  .cw-doc-title { max-width: 120px; }
  .cw-save-text { display: none; }

  /* Outline / Review: overlay the page instead of stealing width on a phone. */
  .cw-outline { position: absolute; left: 0; top: 0; bottom: 0; height: auto; width: min(264px, 80vw); z-index: 30; box-shadow: 2px 0 16px rgba(0,0,0,0.18); }
  .cw-review { position: absolute; right: 0; top: 0; bottom: 0; height: auto; width: min(320px, 88vw); z-index: 31; box-shadow: -2px 0 16px rgba(0,0,0,0.18); }
  .cw-inspector { position: absolute; right: 0; top: 0; bottom: 0; height: auto; width: min(300px, 88vw); z-index: 31; box-shadow: -2px 0 16px rgba(0,0,0,0.18); }

  /* Floating panels (Page Setup, Find) → bottom sheet; drawer fills the pane
     (still inside .cw-workarea, so still below the ribbon).
     !important overrides the inline position/size set in editorApp.ts. */
  .cw-float-panel { left: 8px !important; right: 8px !important; top: auto !important; bottom: 44px !important; width: auto !important; max-width: none !important; max-height: 60vh; overflow: auto; }
  .cw-float-drawer { width: 100%; }
}

/* Touch hit targets: coarse-pointer devices on small screens (phones / small
   tablets). Bounded by width so a wide touchscreen laptop keeps mouse-sized
   controls and the ribbon does not widen into an overflow. */
@media (any-pointer: coarse) and (max-width: 760px) {
  .cw-toolbar button.rib-btn { min-width: 40px; height: 40px; }
  .cw-toolbar button.rib-btn svg { width: 18px; height: 18px; }
  .rib-tab { padding: 9px 14px; font-size: 14px; }
  .cw-toolbar select, .cw-toolbar input[type="number"] { height: 36px; font-size: 14px; }
  .cw-statusbar { height: 36px; }
  .cw-statusbar button.sb-btn { width: 32px; height: 30px; font-size: 17px; }
  .cw-statusbar input[type="range"] { width: 96px; }
}

/* Touch capability at any size: pad the tiny image/shape resize handles so they
   are hittable with a finger. Width-independent and layout-neutral — an
   invisible ::before grows the touch target without changing the visual size. */
@media (any-pointer: coarse) {
  .cw-obj-handle::before { content: ""; position: absolute; inset: -8px; }
}

/* ===== Dark mode — the whole chrome, not just the canvas (critique V2). =====
   Engaged when :root carries data-theme="dark", which editorApp reflects from
   prefers-color-scheme (a host can also set it). The canvas/page is themed
   separately by the DARK_THEME paint preset; this styles the shell around it. */
:root[data-theme="dark"] .cw-app { background: #22252a; }
/* Ribbon */
:root[data-theme="dark"] .cw-toolbar { background: #1e1f22; color: #e6e6e6; }
:root[data-theme="dark"] .rib-tabs { background: #26282c; }
:root[data-theme="dark"] .rib-tab { color: #c8c8c8; }
:root[data-theme="dark"] .rib-tab:hover { background: #34373c; }
:root[data-theme="dark"] .rib-tab.active { background: #1e1f22; color: #7cb0ff; box-shadow: none; }
:root[data-theme="dark"] .rib-tab.active::after { background: #7cb0ff; }
:root[data-theme="dark"] .rib-tab.rib-tab-ctx { color: #c58af9; }
:root[data-theme="dark"] .rib-tab.rib-tab-ctx.active { color: #c58af9; }
:root[data-theme="dark"] .rib-tab.rib-tab-ctx.active::after { background: #c58af9; }
:root[data-theme="dark"] .rib-bodies, :root[data-theme="dark"] .rib-panel { background: #1e1f22; }
:root[data-theme="dark"] .rib-group { border-color: #34373c; }
:root[data-theme="dark"] .rib-label { color: #9aa0a6; }
:root[data-theme="dark"] .rib-tab-overflow { color: #c8c8c8; }
:root[data-theme="dark"] .rib-tab-overflow:hover { background: #34373c; }
:root[data-theme="dark"] .cw-toolbar button.rib-btn { color: #e0e0e0; }
:root[data-theme="dark"] .cw-toolbar button.rib-btn:hover { background: #34373c; }
:root[data-theme="dark"] .cw-toolbar button.rib-btn:active { background: #3c4046; }
:root[data-theme="dark"] .cw-toolbar button.rib-btn.active { background: #2a3f5f; color: #7cb0ff; border-color: #3d6ba5; }
:root[data-theme="dark"] .cw-toolbar button.rib-btn.active svg { color: #7cb0ff; }
:root[data-theme="dark"] .cw-toolbar button.rib-btn:disabled { color: #6a6f76; }
:root[data-theme="dark"] .cw-toolbar select,
:root[data-theme="dark"] .cw-toolbar input[type="number"] { background: #2a2c30; color: #e6e6e6; border-color: #4a4a4a; }
/* Doc identity + save state + quick-access + header controls */
:root[data-theme="dark"] .cw-doc-title { color: #f0f0f0; }
:root[data-theme="dark"] .cw-save-text { color: #9aa0a6; }
:root[data-theme="dark"] .cw-qat { border-color: #34373c; }
:root[data-theme="dark"] .cw-qat-btn { color: #e0e0e0; }
:root[data-theme="dark"] .cw-qat-btn:hover:not(:disabled) { background: #34373c; }
:root[data-theme="dark"] .cw-mode-select { background: #2a2c30; color: #e6e6e6; border-color: #4a4a4a; }
:root[data-theme="dark"] .cw-header-btn { background: #2a2c30; color: #7cb0ff; border-color: #4a4a4a; }
:root[data-theme="dark"] .cw-header-btn:hover { background: #34373c; }
:root[data-theme="dark"] .cw-fidelity.faithful { color: #8fd6a4; background: #1a2b1f; border-color: #2f5138; }
:root[data-theme="dark"] .cw-fidelity.notes { color: #e6bd73; background: #2e2517; border-color: #5a4a2a; }
:root[data-theme="dark"] .cw-fidelity-panel { background: #26282c; border-color: #3a3d42; color: #e6e6e6; }
:root[data-theme="dark"] .cw-fidelity-sub { color: #9aa0a6; }
:root[data-theme="dark"] .cw-fidelity-list li { color: #c8ccd0; }
/* Styles gallery */
:root[data-theme="dark"] .rib-gallery { background: #2a2c30; border-color: #4a4a4a; }
:root[data-theme="dark"] .style-card { background: #26282c; border-color: #3a3a3a; }
:root[data-theme="dark"] .style-card .name { color: #9aa0a6; }
:root[data-theme="dark"] .style-card .preview { color: #e6e6e6; }
/* Status bar */
:root[data-theme="dark"] .cw-statusbar { background: #17181b; color: #d0d0d0; }
/* Outline pane */
:root[data-theme="dark"] .cw-outline { background: #1e1f22; border-color: #34373c; }
:root[data-theme="dark"] .cw-outline .outline-head { background: #1e1f22; color: #e6e6e6; border-color: #34373c; }
:root[data-theme="dark"] .cw-outline .outline-head button { color: #9aa0a6; }
:root[data-theme="dark"] .cw-outline .outline-head button:hover { background: #34373c; }
:root[data-theme="dark"] .cw-outline-filter { border-color: #2c2f34; }
:root[data-theme="dark"] .cw-outline-filter input { background: #2a2c30; color: #e6e6e6; border-color: #4a4a4a; }
:root[data-theme="dark"] .outline-item { color: #d0d0d0; }
:root[data-theme="dark"] .outline-item:hover { background: #2c2f34; }
:root[data-theme="dark"] .outline-item.active { background: #2a3f5f; border-left-color: #7cb0ff; }
:root[data-theme="dark"] .outline-item.active .outline-label { color: #7cb0ff; }
:root[data-theme="dark"] .outline-item.lvl-0 .outline-label,
:root[data-theme="dark"] .outline-item.lvl-1 .outline-label { color: #f0f0f0; }
:root[data-theme="dark"] .outline-item.lvl-2 .outline-label { color: #c8c8c8; }
:root[data-theme="dark"] .outline-item.lvl-3 .outline-label { color: #9aa0a6; }
:root[data-theme="dark"] .outline-page { color: #6b7280; }
:root[data-theme="dark"] .outline-chevron { color: #80868b; }
:root[data-theme="dark"] .outline-chevron:hover { background: #34373c; color: #e6e6e6; }
:root[data-theme="dark"] .outline-empty { color: #80868b; }
/* Review pane + right-docked drawers */
:root[data-theme="dark"] .cw-review { background: #1e1f22; border-color: #34373c; color: #d0d0d0; }
:root[data-theme="dark"] .cw-review-head, :root[data-theme="dark"] .cw-review-tabs { background: #1e1f22; border-color: #34373c; }
:root[data-theme="dark"] .cw-float-drawer { background: #1e1f22; border-color: #34373c; color: #d0d0d0; }
/* Shared popovers / menus / dialogs */
:root[data-theme="dark"] .cw-pop { background: #26282c; border-color: #4a4a4a; color: #e6e6e6; }
:root[data-theme="dark"] .cw-menu { background: #26282c; border-color: #4a4a4a; color: #e6e6e6; }
:root[data-theme="dark"] .cw-dialog input,
:root[data-theme="dark"] .cw-dialog button { background: #2a2c30; color: #e6e6e6; border-color: #4a4a4a; }
:root[data-theme="dark"] .cw-dialog button.primary { background: #3d6ba5; border-color: #3d6ba5; color: #fff; }
/* Every dialogShell dialog carries role=dialog on its modal — theme generically. */
:root[data-theme="dark"] [role="dialog"] { background: #26282c; color: #e6e6e6; border-color: #4a4a4a; }
:root[data-theme="dark"] [role="dialog"] h2 { color: #f0f0f0; }
:root[data-theme="dark"] [role="dialog"] input,
:root[data-theme="dark"] [role="dialog"] select,
:root[data-theme="dark"] [role="dialog"] textarea { background: #1e1f22; color: #e6e6e6; border-color: #4a4a4a; }
/* Busy overlay card */
:root[data-theme="dark"] .cw-loading-card { background: #26282c; border-color: #4a4a4a; color: #e6e6e6; }
/* Find bar — its box/inputs/buttons are inline-styled, so override with !important. */
:root[data-theme="dark"] .cw-float-panel { background: #26282c !important; border-color: #4a4a4a !important; color: #e6e6e6; }
:root[data-theme="dark"] .cw-float-panel input { background: #1e1f22 !important; color: #e6e6e6 !important; border-color: #4a4a4a !important; }
:root[data-theme="dark"] .cw-float-panel button { color: #d0d0d0 !important; }
`;

/** Append a <style id> with the given css once per document, keyed by `id`
 *  (idempotent across calls and hot-reloads). Shared by the editor chrome and
 *  the on-demand modal/menu components so the inject-once pattern lives once. */
export function injectCssOnce(id: string, css: string): void {
  if (document.getElementById(id)) return;
  const style = document.createElement("style");
  style.id = id;
  style.textContent = css;
  document.head.appendChild(style);
}

/** Inject the component stylesheet once per document (idempotent). */
export function ensureWordCanvasStyles(): void {
  injectCssOnce(STYLE_ID, CSS);
}
