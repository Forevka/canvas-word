// [Content_Types].xml → locate the main document part. Usually
// "word/document.xml", but the spec allows any part name (and macro-enabled
// .docm uses a different content type), so don't hardcode.

import { attr, children, parseXml, rootEl } from "./xml";

const MAIN_CONTENT_TYPES = new Set([
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  "application/vnd.ms-word.document.macroEnabled.main+xml",
]);

export function findMainDocumentPart(contentTypesXml: string): string | undefined {
  const types = rootEl(parseXml(contentTypesXml, "[Content_Types].xml"), "Types");
  if (!types) return undefined;
  for (const child of children(types)) {
    if (child.tagName !== "Override") continue;
    const ct = attr(child, "ContentType");
    const partName = attr(child, "PartName");
    if (ct && partName && MAIN_CONTENT_TYPES.has(ct)) {
      return partName.replace(/^\//, ""); // OPC part names are /-rooted; zip entries aren't
    }
  }
  return undefined;
}
