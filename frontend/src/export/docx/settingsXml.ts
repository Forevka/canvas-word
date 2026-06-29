// settings.xml — the document-level settings the model round-trips: distinct
// even/odd headers/footers, background-shape display (so Word paints the page
// color), a non-default tab interval (w:defaultTabStop), and any w:compat flags.

import { pxToTwips } from "../units";
import { el, WML_NS, XML_DECL } from "./xmlWrite";

export interface SettingsXmlOptions {
  evenAndOdd: boolean;
  displayBackgroundShape?: boolean;
  /** Default tab interval in px → emitted as w:defaultTabStop (twips). */
  defaultTabStopPx?: number;
  /** w:compat/w:compatSetting triples to re-emit verbatim. */
  compatSettings?: { name: string; uri: string; val: string }[];
}

export function settingsXml(opts: SettingsXmlOptions): string {
  // w:displayBackgroundShape makes Word actually paint w:background (the page color).
  const compat =
    opts.compatSettings && opts.compatSettings.length > 0
      ? el(
          "w:compat",
          undefined,
          opts.compatSettings
            .map((c) => el("w:compatSetting", { "w:name": c.name, "w:uri": c.uri, "w:val": c.val }))
            .join(""),
        )
      : "";
  const body =
    (opts.displayBackgroundShape ? el("w:displayBackgroundShape") : "") +
    (opts.evenAndOdd ? el("w:evenAndOddHeadersAndFooters") : "") +
    (opts.defaultTabStopPx !== undefined ? el("w:defaultTabStop", { "w:val": Math.round(pxToTwips(opts.defaultTabStopPx)) }) : "") +
    compat;
  return XML_DECL + el("w:settings", WML_NS, body);
}
