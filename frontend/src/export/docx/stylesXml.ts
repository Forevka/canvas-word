// styles.xml from Document.stylesheet — paragraph and character styles with their
// OWN deltas (the editor resolves basedOn chains). Character styles emit
// w:type="character" and carry only rPr (no pPr). Inverse of
// mapToModel.buildStylesheet.

import { styleType, type Stylesheet, type TableStyle } from "@cw/shared";
import { partialPPrXml, partialRPrXml, tableStyleXml } from "./styleProps";
import { el, WML_NS, XML_DECL } from "./xmlWrite";

export function stylesXml(sheet: Stylesheet | undefined, tableStyles?: Record<string, TableStyle>): string {
  const named = (sheet?.styles ?? [])
    .map((s) => {
      const isChar = styleType(s) === "character";
      const attrs: Record<string, string> = { "w:type": isChar ? "character" : "paragraph", "w:styleId": s.id };
      if (!isChar && sheet && s.id === sheet.defaultStyleId) attrs["w:default"] = "1";
      const body =
        el("w:name", { "w:val": s.name }) +
        (s.basedOn ? el("w:basedOn", { "w:val": s.basedOn }) : "") +
        (isChar ? "" : partialPPrXml(s.para)) +
        partialRPrXml(s.char);
      return el("w:style", attrs, body);
    })
    .join("");
  const tables = Object.values(tableStyles ?? {}).map(tableStyleXml).join("");
  return XML_DECL + el("w:styles", WML_NS, named + tables);
}
