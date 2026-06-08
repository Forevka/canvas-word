// settings.xml — minimal. Enables distinct even/odd headers/footers when the
// document carries even-page bands (the importer reads this to gate them).

import { el, WML_NS, XML_DECL } from "./xmlWrite";

export function settingsXml(evenAndOdd: boolean): string {
  const body = evenAndOdd ? el("w:evenAndOddHeadersAndFooters") : "";
  return XML_DECL + el("w:settings", WML_NS, body);
}
