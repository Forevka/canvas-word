// Minimal, dependency-free XML string emitter for OOXML parts. Attribute order is
// insertion order (Word doesn't care); empty children collapse to self-closing.

export type Attrs = Record<string, string | number | boolean | undefined>;

export function escapeText(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeAttr(s: string): string {
  return escapeText(s).replace(/"/g, "&quot;");
}

function serializeAttrs(attrs: Attrs | undefined): string {
  if (!attrs) return "";
  let out = "";
  for (const [k, v] of Object.entries(attrs)) {
    if (v === undefined || v === false) continue;
    out += ` ${k}="${escapeAttr(String(v))}"`;
  }
  return out;
}

/** An element. `children` is pre-serialized XML (caller joins arrays). Pass ""
 *  or undefined for an empty element (self-closing). */
export function el(tag: string, attrs?: Attrs, children?: string): string {
  const a = serializeAttrs(attrs);
  if (children === undefined || children === "") return `<${tag}${a}/>`;
  return `<${tag}${a}>${children}</${tag}>`;
}

/** A w:t, preserving leading/trailing whitespace when present. */
export function textEl(text: string): string {
  const needsPreserve = /^\s|\s$|\s\s/.test(text) || text.includes("\t");
  return el("w:t", needsPreserve ? { "xml:space": "preserve" } : undefined, escapeText(text));
}

export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/** The w: + relationships + drawing namespaces Word expects on a document root. */
export const WML_NS = {
  "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "xmlns:wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
  "xmlns:pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
  "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
  "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
  "xmlns:m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
} as const;
