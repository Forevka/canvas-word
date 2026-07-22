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
.cw-workarea { flex: 1 1 auto; min-height: 0; display: flex; }
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
.cw-outline {
  flex: 0 0 264px; width: 264px; min-height: 0; overflow-y: auto;
  background: #fff; border-right: 1px solid #e1dfdd; display: none;
}
.cw-outline.open { display: block; }
.cw-outline .outline-head {
  position: sticky; top: 0; background: #fff; z-index: 1;
  display: flex; align-items: center; justify-content: space-between;
  padding: 9px 8px 9px 14px; border-bottom: 1px solid #e1dfdd;
  font-size: 13px; font-weight: 600; color: #323130;
}
.cw-outline .outline-head button {
  border: none; background: transparent; cursor: pointer; color: #605e5c;
  width: 24px; height: 24px; border-radius: 4px; font-size: 17px; line-height: 1;
}
.cw-outline .outline-head button:hover { background: #f3f2f1; }
.cw-outline-list { padding: 4px 0 12px; }
.outline-item {
  display: block; width: 100%; box-sizing: border-box; text-align: left;
  border: none; border-left: 3px solid transparent; background: transparent; cursor: pointer;
  padding: 4px 12px; color: #323130; font-size: 13px; line-height: 1.35;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}
.outline-item:hover { background: #f3f2f1; }
.outline-item.active { background: #eef3fb; border-left-color: #2b579a; color: #2b579a; font-weight: 600; }
.outline-empty { padding: 14px; color: #80868b; font-size: 12px; line-height: 1.4; }

/* ===== Review pane (track changes + comments) ======================= */
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
  .cw-toolbar button.rib-btn { min-width: 40px; height: 40px; }
  .cw-toolbar button.rib-btn svg { width: 18px; height: 18px; }
  .rib-tab { padding: 9px 14px; font-size: 14px; }
  .cw-toolbar select, .cw-toolbar input[type="number"] { height: 36px; font-size: 14px; }

  /* Outline: overlay the page instead of stealing 264px of a ~360px screen. */
  .cw-workarea { position: relative; }
  .cw-outline { position: absolute; left: 0; top: 0; bottom: 0; height: auto; width: min(264px, 80vw); z-index: 30; box-shadow: 2px 0 16px rgba(0,0,0,0.18); }
  .cw-review { position: absolute; right: 0; top: 0; bottom: 0; height: auto; width: min(320px, 88vw); z-index: 31; box-shadow: -2px 0 16px rgba(0,0,0,0.18); }

  /* Status bar: tappable zoom controls. */
  .cw-statusbar { height: 36px; }
  .cw-statusbar button.sb-btn { width: 32px; height: 30px; font-size: 17px; }
  .cw-statusbar input[type="range"] { width: 96px; }

  /* Floating panels (Page Setup, Find) → bottom sheet; Activity → full-width.
     !important overrides the inline position/size set in editorApp.ts. */
  .cw-float-panel { left: 8px !important; right: 8px !important; top: auto !important; bottom: 8px !important; width: auto !important; max-width: none !important; max-height: 60vh; overflow: auto; }
  .cw-float-drawer { width: 100% !important; }

  /* Image resize handles: 8px dots are unhittable with a finger — an invisible
     ::before pads the touch target to ~24px without changing the visual size. */
  .cw-obj-handle::before { content: ""; position: absolute; inset: -8px; }
}
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
