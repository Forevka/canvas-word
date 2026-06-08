// Ribbon icon set — hand-authored 16×16 SVGs (Fluent-like single-color
// strokes/fills on currentColor) so the toolbar carries no emoji glyphs.
// Pure data: each entry is an innerHTML-ready <svg> string.

const svg = (body: string): string =>
  `<svg viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round">${body}</svg>`;

const filled = (body: string): string =>
  `<svg viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg" fill="currentColor">${body}</svg>`;

export const ICONS = {
  undo: svg(`<path d="M5.5 3.5 2.5 6.5l3 3"/><path d="M2.5 6.5h7a4 4 0 0 1 0 8H6"/>`),
  redo: svg(`<path d="M10.5 3.5l3 3-3 3"/><path d="M13.5 6.5h-7a4 4 0 0 0 0 8H10"/>`),
  painter: svg(
    `<path d="M3 2.5h8v3H3z"/><path d="M11 4h2.5v3.5H7V10"/><path d="M6 10h2v4H6z"/>`,
  ),
  open: svg(`<path d="M1.5 4.5v8h11l2-5H4l-1.5 5"/><path d="M1.5 4.5l1-2h4l1 2h5v2"/>`),
  save: svg(`<path d="M2.5 2.5h9l2 2v9h-11z"/><path d="M5 2.5v3.5h5V2.5"/><path d="M4.5 13.5V9h7v4.5"/>`),
  link: svg(
    `<path d="M6.5 9.5 9.5 6.5"/><path d="M7.5 4.5 9 3a2.8 2.8 0 0 1 4 4l-1.5 1.5"/><path d="M8.5 11.5 7 13a2.8 2.8 0 0 1-4-4l1.5-1.5"/>`,
  ),
  highlight: svg(
    `<path d="M3 13.5h4"/><path d="M5.5 11 11 4.5l2 2L7.5 13l-2.5.5z"/><path d="M10 5.5l2 2"/>`,
  ),
  bullets: filled(
    `<circle cx="3" cy="4" r="1.2"/><circle cx="3" cy="8" r="1.2"/><circle cx="3" cy="12" r="1.2"/><rect x="6" y="3.4" width="8" height="1.2" rx=".6"/><rect x="6" y="7.4" width="8" height="1.2" rx=".6"/><rect x="6" y="11.4" width="8" height="1.2" rx=".6"/>`,
  ),
  numbering: filled(
    `<text x="1.4" y="5.4" font-size="5" font-family="Arial" stroke="none">1</text><text x="1.4" y="9.6" font-size="5" font-family="Arial" stroke="none">2</text><text x="1.4" y="13.8" font-size="5" font-family="Arial" stroke="none">3</text><rect x="6" y="3.4" width="8" height="1.2" rx=".6"/><rect x="6" y="7.6" width="8" height="1.2" rx=".6"/><rect x="6" y="11.8" width="8" height="1.2" rx=".6"/>`,
  ),
  alignLeft: filled(
    `<rect x="2" y="3" width="12" height="1.2" rx=".6"/><rect x="2" y="6" width="8" height="1.2" rx=".6"/><rect x="2" y="9" width="12" height="1.2" rx=".6"/><rect x="2" y="12" width="8" height="1.2" rx=".6"/>`,
  ),
  alignCenter: filled(
    `<rect x="2" y="3" width="12" height="1.2" rx=".6"/><rect x="4" y="6" width="8" height="1.2" rx=".6"/><rect x="2" y="9" width="12" height="1.2" rx=".6"/><rect x="4" y="12" width="8" height="1.2" rx=".6"/>`,
  ),
  alignRight: filled(
    `<rect x="2" y="3" width="12" height="1.2" rx=".6"/><rect x="6" y="6" width="8" height="1.2" rx=".6"/><rect x="2" y="9" width="12" height="1.2" rx=".6"/><rect x="6" y="12" width="8" height="1.2" rx=".6"/>`,
  ),
  alignJustify: filled(
    `<rect x="2" y="3" width="12" height="1.2" rx=".6"/><rect x="2" y="6" width="12" height="1.2" rx=".6"/><rect x="2" y="9" width="12" height="1.2" rx=".6"/><rect x="2" y="12" width="12" height="1.2" rx=".6"/>`,
  ),
  lineSpacing: svg(
    `<path d="M3.5 3v10"/><path d="M2 4.5 3.5 3 5 4.5"/><path d="M2 11.5 3.5 13 5 11.5"/><path d="M8 4h6"/><path d="M8 8h6"/><path d="M8 12h6"/>`,
  ),
  image: svg(
    `<rect x="2" y="3" width="12" height="10" rx="1"/><circle cx="5.5" cy="6.5" r="1.2"/><path d="M2 11l3.5-3 3 2.5L11 8l3 3"/>`,
  ),
  table: svg(
    `<rect x="2" y="2.5" width="12" height="11" rx="1"/><path d="M2 6h12"/><path d="M2 9.7h12"/><path d="M6.5 2.5v11"/><path d="M10.7 2.5v11"/>`,
  ),
  pageBreak: svg(
    `<path d="M3 2.5h10v4H3z"/><path d="M3 13.5h10v-4H3z" /><path d="M2 8h2.4M6 8h1.6M9.6 8h1.6M13 8h1"/>`,
  ),
  sectionBreak: svg(
    `<path d="M3 4h10"/><path d="M3 12h10"/><path d="M2 8h2.4M6 8h1.6M9.6 8h1.6M13 8h1"/><path d="M6 1.5h4M6 14.5h4"/>`,
  ),
  columns: svg(
    `<rect x="2" y="3" width="5" height="10" rx=".8"/><rect x="9" y="3" width="5" height="10" rx=".8"/>`,
  ),
  toc: filled(
    `<rect x="2" y="3" width="7" height="1.2" rx=".6"/><rect x="11.5" y="3" width="2.5" height="1.2" rx=".6"/><rect x="3.5" y="7" width="5.5" height="1.2" rx=".6"/><rect x="11.5" y="7" width="2.5" height="1.2" rx=".6"/><rect x="2" y="11" width="7" height="1.2" rx=".6"/><rect x="11.5" y="11" width="2.5" height="1.2" rx=".6"/>`,
  ),
  pageSetup: svg(
    `<rect x="3" y="2" width="10" height="12" rx="1"/><path d="M3 5h10"/><path d="M5.5 8h5M5.5 10.5h5"/>`,
  ),
  rowAbove: svg(
    `<rect x="2" y="8" width="12" height="6" rx="1"/><path d="M8 8v6M2 11h12"/><path d="M8 5.5V1.5"/><path d="M6 3.5 8 1.5l2 2"/>`,
  ),
  rowBelow: svg(
    `<rect x="2" y="2" width="12" height="6" rx="1"/><path d="M8 2v6M2 5h12"/><path d="M8 10.5v4"/><path d="M6 12.5l2 2 2-2"/>`,
  ),
  colLeft: svg(
    `<rect x="8" y="2" width="6" height="12" rx="1"/><path d="M8 8h6M11 2v12"/><path d="M5.5 8h-4"/><path d="M3.5 6 1.5 8l2 2"/>`,
  ),
  colRight: svg(
    `<rect x="2" y="2" width="6" height="12" rx="1"/><path d="M2 8h6M5 2v12"/><path d="M10.5 8h4"/><path d="M12.5 6l2 2-2 2"/>`,
  ),
  deleteRow: svg(
    `<rect x="2" y="5" width="12" height="6" rx="1"/><path d="M6.2 5v6M9.8 5v6"/><path d="M11 1.5l3 3M14 1.5l-3 3"/>`,
  ),
  deleteCol: svg(
    `<rect x="5" y="2" width="6" height="12" rx="1"/><path d="M5 6.2h6M5 9.8h6"/><path d="M11.5 11.5l3 3M14.5 11.5l-3 3"/>`,
  ),
  deleteTable: svg(
    `<rect x="2" y="2.5" width="12" height="11" rx="1"/><path d="M2 6h12M6.5 2.5v11"/><path d="M9 9l4 4M13 9l-4 4"/>`,
  ),
  mergeCells: svg(
    `<rect x="2" y="3" width="12" height="10" rx="1"/><path d="M5.5 8h5"/><path d="M8.5 6.5 10.5 8l-2 1.5"/><path d="M2 3v10M14 3v10"/>`,
  ),
  unmergeCells: svg(
    `<rect x="2" y="3" width="12" height="10" rx="1"/><path d="M8 3v10"/><path d="M6.5 6.5 4.5 8l2 1.5"/><path d="M9.5 6.5 11.5 8l-2 1.5"/>`,
  ),
  wrapSquare: svg(
    `<rect x="2" y="5" width="5" height="5" rx=".8"/><path d="M2 2.5h12M9 5.5h5M9 8h5M2 12.5h12"/>`,
  ),
  wrapInline: svg(
    `<rect x="5" y="5" width="6" height="5" rx=".8"/><path d="M2 2.5h12M2 12.5h12"/>`,
  ),
  sdtText: svg(
    `<rect x="1.5" y="3.5" width="13" height="9" rx="1.5"/><path d="M5 6.2h6"/><path d="M8 6.2v4"/>`,
  ),
  sdtCheckbox: svg(
    `<rect x="1.5" y="3.5" width="13" height="9" rx="1.5"/><path d="M5.5 8l1.8 1.8L10.8 6"/>`,
  ),
  sdtDropdown: svg(
    `<rect x="1.5" y="3.5" width="13" height="9" rx="1.5"/><path d="M10.5 3.5v9"/><path d="M11.6 7l1.2 1.4L14 7"/><path d="M3.5 7h4.5"/>`,
  ),
  sdtDate: svg(
    `<rect x="2" y="3" width="12" height="11" rx="1.5"/><path d="M2 6.5h12"/><path d="M5 1.8v2.4M11 1.8v2.4"/><path d="M4.5 9.5h2M9.5 9.5h2M4.5 12h2"/>`,
  ),
  sdtRemove: svg(
    `<rect x="1.5" y="3.5" width="13" height="9" rx="1.5" stroke-dasharray="2.5 1.8"/><path d="M6 6l4 4M10 6l-4 4"/>`,
  ),
  sdtProps: svg(
    `<rect x="1.5" y="3.5" width="13" height="9" rx="1.5"/><path d="M8 6.2v.2"/><path d="M8 8v2.3"/>`,
  ),
  find: svg(`<circle cx="7" cy="7" r="4.5"/><path d="M10.5 10.5 14 14"/>`),
  stylePencil: svg(`<path d="M3 13l.8-3L10.5 3.3a1.5 1.5 0 0 1 2.2 2.2L6 12.2 3 13z"/><path d="M9.5 4.5l2 2"/>`),
  styleNew: svg(`<path d="M8 3v10M3 8h10"/>`),
} as const;

export type IconName = keyof typeof ICONS;
