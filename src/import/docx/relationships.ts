// OPC relationships: word/_rels/document.xml.rels maps rIds → part targets.
// Used to locate styles.xml / theme (milestone 2) and media parts (milestone 3)
// instead of hardcoding conventional names.

import { attr, children, parseXml, rootEl } from "./xml";

export interface Relationship {
  id: string;
  /** Full relationship type URI; match with findByType() on the trailing segment. */
  type: string;
  /** Internal: resolved zip part name. External: the raw target (URL/path) —
   *  hyperlinks always, but also linked ("Link to File") images and OLE objects. */
  target: string;
  /** TargetMode="External" — target lives outside the package. */
  external: boolean;
}

export type Relationships = Map<string, Relationship>;

/** @param sourcePart the part these rels belong to, e.g. "word/document.xml" —
 *  relative targets resolve against its directory. */
export function parseRelationships(relsXml: string, sourcePart: string): Relationships {
  const baseDir = sourcePart.includes("/") ? sourcePart.slice(0, sourcePart.lastIndexOf("/")) : "";
  const out: Relationships = new Map();
  const root = rootEl(parseXml(relsXml, `${sourcePart} rels`), "Relationships");
  if (!root) return out;
  for (const rel of children(root)) {
    if (rel.tagName !== "Relationship") continue;
    const id = attr(rel, "Id");
    const type = attr(rel, "Type");
    const target = attr(rel, "Target");
    if (!id || !type || !target) continue;
    const external = attr(rel, "TargetMode") === "External";
    out.set(id, { id, type, target: external ? target : resolveTarget(baseDir, target), external });
  }
  return out;
}

/** "/word/styles.xml" is package-rooted; "styles.xml" or "../x.xml" is relative. */
function resolveTarget(baseDir: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = baseDir === "" ? [] : baseDir.split("/");
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== "." && seg !== "") parts.push(seg);
  }
  return parts.join("/");
}

/** Relationship types share a URI prefix; the trailing segment is the kind. */
export function findByType(rels: Relationships, kind: string): Relationship | undefined {
  for (const rel of rels.values()) {
    if (rel.type.endsWith(`/${kind}`)) return rel;
  }
  return undefined;
}

/** Conventional location of a part's own rels: dir/_rels/name.rels */
export function relsPartFor(partName: string): string {
  const slash = partName.lastIndexOf("/");
  const dir = slash >= 0 ? partName.slice(0, slash + 1) : "";
  const name = slash >= 0 ? partName.slice(slash + 1) : partName;
  return `${dir}_rels/${name}.rels`;
}
