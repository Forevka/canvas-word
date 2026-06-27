// Tiny txml-node helpers shared by the MathML parser and the OMML importer.
// Both match elements by LOCAL name (prefix-insensitive), so `mi`, `m:f`,
// `mml:mrow` all work regardless of how the producer declared namespaces.

import { isElementNode, type TNode } from "txml/txml";

/** Local (unprefixed) name of a qualified tag. */
export const local = (tag: string): string => {
  const i = tag.indexOf(":");
  return i === -1 ? tag : tag.slice(i + 1);
};

/** Element children only (text nodes filtered out). */
export const childEls = (n: TNode): TNode[] => n.children.filter(isElementNode);

/** First child element with the given local name. */
export const childByLocal = (n: TNode, name: string): TNode | undefined =>
  childEls(n).find((c) => local(c.tagName) === name);

/** Concatenated direct text content, trimmed. */
export const text = (n: TNode): string => {
  let out = "";
  for (const c of n.children) if (typeof c === "string") out += c;
  return out.trim();
};

/** Attribute by qualified name; valueless attributes read as "". */
export const attr = (n: TNode, name: string): string | undefined => {
  const v = n.attributes[name];
  return v === undefined ? undefined : (v ?? "");
};

/** OMML `<m:x m:val="…"/>` child-property reader (prefix-insensitive on the
 *  child name; the val attribute may be `m:val` or `w:val`). */
export const mathVal = (n: TNode, childLocal: string): string | undefined => {
  const c = childByLocal(n, childLocal);
  if (!c) return undefined;
  return attr(c, "m:val") ?? attr(c, "w:val") ?? "";
};

/** Is an OMML on/off property present and not explicitly off? */
export const mathOn = (n: TNode, childLocal: string): boolean => {
  const c = childByLocal(n, childLocal);
  if (!c) return false;
  const v = attr(c, "m:val") ?? attr(c, "w:val");
  return v === undefined || !["0", "false", "off"].includes(v);
};
