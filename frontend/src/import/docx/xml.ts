// txml helpers. DOMParser is window-only (not available in workers), so we use
// txml — a tiny pure-JS parser that returns an ORDERED node tree. Order matters:
// in document.xml, run order IS the text.
//
// OOXML namespace handling: prefixes (w:, a:, wp:) are technically remappable,
// but Word always emits the conventional ones, so we match qualified names
// directly — the same pragmatic choice docx-preview and mammoth make.

import { isElementNode, parse, type TNode } from "txml/txml";
import { ImportError } from "./types";

export type XmlNode = TNode;

/** Parse an XML part. keepWhitespace matters: a run whose text is a single
 *  space (`<w:t xml:space="preserve"> </w:t>`) must not be dropped. */
export function parseXml(text: string, partName: string): XmlNode[] {
  try {
    return parse(text, { keepWhitespace: true, decodeEntities: true }).filter(isElementNode);
  } catch (e) {
    throw new ImportError(
      "MALFORMED_XML",
      `Failed to parse ${partName}: ${e instanceof Error ? e.message : String(e)}`,
    );
  }
}

/** Top-level element by qualified name (skips the XML declaration / PIs). */
export function rootEl(nodes: XmlNode[], tagName: string): XmlNode | undefined {
  return nodes.find((n) => n.tagName === tagName);
}

/** Element children only (text nodes filtered out). */
export function children(node: XmlNode): XmlNode[] {
  return node.children.filter(isElementNode);
}

/** First child element with the given qualified name. */
export function el(node: XmlNode, tagName: string): XmlNode | undefined {
  return node.children.find((c): c is XmlNode => isElementNode(c) && c.tagName === tagName);
}

/** All child elements with the given qualified name. */
export function els(node: XmlNode, tagName: string): XmlNode[] {
  return node.children.filter((c): c is XmlNode => isElementNode(c) && c.tagName === tagName);
}

/** Attribute value by qualified name; valueless attributes read as "". */
export function attr(node: XmlNode, name: string): string | undefined {
  const v = node.attributes[name];
  return v === undefined ? undefined : (v ?? "");
}

/** Depth-first search for a descendant element. Used where the exact nesting
 *  is deep and producer-dependent (e.g. a:blip inside w:drawing). */
export function findDeep(node: XmlNode, tagName: string): XmlNode | undefined {
  for (const c of node.children) {
    if (!isElementNode(c)) continue;
    if (c.tagName === tagName) return c;
    const found = findDeep(c, tagName);
    if (found) return found;
  }
  return undefined;
}

/** Concatenated text content (direct string children — w:t never nests). */
export function textOf(node: XmlNode): string {
  let out = "";
  for (const c of node.children) if (typeof c === "string") out += c;
  return out;
}

/** Shorthand for the ubiquitous `<w:x w:val="…"/>` child-element pattern. */
export function val(node: XmlNode, tagName: string): string | undefined {
  const child = el(node, tagName);
  return child ? attr(child, "w:val") : undefined;
}

/** OOXML on/off property: `<w:b/>` is on; w:val of 0/false/none/off is off. */
export function onOff(node: XmlNode | undefined): boolean | undefined {
  if (!node) return undefined;
  const v = attr(node, "w:val");
  if (v === undefined) return true;
  return !["0", "false", "none", "off"].includes(v);
}

/** Numeric attribute (OOXML writes plain decimal integers). */
export function numAttr(node: XmlNode | undefined, name: string): number | undefined {
  if (!node) return undefined;
  const v = attr(node, name);
  if (v === undefined) return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}
