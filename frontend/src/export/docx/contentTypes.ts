// [Content_Types].xml — default extension rules plus one Override per part.

import { el, XML_DECL } from "./xmlWrite";

const CT = {
  document: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  styles: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
  numbering: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
  footnotes: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
  endnotes: "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
  settings: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
  header: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
  footer: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
  theme: "application/vnd.openxmlformats-officedocument.theme+xml",
} as const;

export { CT };

export function contentTypesXml(overrides: [string, string][], imageExts: [string, string][]): string {
  const defaults =
    el("Default", { Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml" }) +
    el("Default", { Extension: "xml", ContentType: "application/xml" }) +
    imageExts.map(([ext, ct]) => el("Default", { Extension: ext, ContentType: ct })).join("");
  const over = overrides.map(([part, ct]) => el("Override", { PartName: part, ContentType: ct })).join("");
  return (
    XML_DECL +
    el("Types", { xmlns: "http://schemas.openxmlformats.org/package/2006/content-types" }, defaults + over)
  );
}
