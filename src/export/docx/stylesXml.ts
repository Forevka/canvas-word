// styles.xml from Document.stylesheet — paragraph styles with their OWN deltas
// (the editor resolves basedOn chains). Inverse of mapToModel.buildStylesheet.

import type { Stylesheet } from "../../model/stylesheet";
import { partialPPrXml, partialRPrXml } from "./styleProps";
import { el, WML_NS, XML_DECL } from "./xmlWrite";

export function stylesXml(sheet: Stylesheet): string {
  const styles = sheet.styles
    .map((s) => {
      const attrs: Record<string, string> = { "w:type": "paragraph", "w:styleId": s.id };
      if (s.id === sheet.defaultStyleId) attrs["w:default"] = "1";
      const body =
        el("w:name", { "w:val": s.name }) +
        (s.basedOn ? el("w:basedOn", { "w:val": s.basedOn }) : "") +
        partialPPrXml(s.para) +
        partialRPrXml(s.char);
      return el("w:style", attrs, body);
    })
    .join("");
  return XML_DECL + el("w:styles", WML_NS, styles);
}
