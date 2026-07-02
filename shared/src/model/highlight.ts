// Word's 16 named highlight colors — the OOXML w:highlight/@w:val vocabulary and
// the hex values the model stores in CharStyle.highlightColor. Single source for
// the docx importer (name → hex) and exporter (hex → name), which previously
// hand-maintained the two maps as mutual inverses (adding a color to one and not
// the other silently lost round-trip fidelity for it).

/** w:highlight name → the hex the model stores. */
export const HIGHLIGHT_HEX: Record<string, string> = {
  yellow: "#ffff00", green: "#00ff00", cyan: "#00ffff", magenta: "#ff00ff",
  blue: "#0000ff", red: "#ff0000", darkBlue: "#000080", darkCyan: "#008080",
  darkGreen: "#008000", darkMagenta: "#800080", darkRed: "#800000", darkYellow: "#808000",
  darkGray: "#808080", lightGray: "#c0c0c0", black: "#000000", white: "#ffffff",
};

/** model hex → w:highlight name (the derived inverse — never hand-edit). */
export const HIGHLIGHT_NAME: Record<string, string> = Object.fromEntries(
  Object.entries(HIGHLIGHT_HEX).map(([name, hex]) => [hex, name]),
);
