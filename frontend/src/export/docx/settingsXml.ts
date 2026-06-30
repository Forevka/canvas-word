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
  // Mirror parseSettings' acceptance rules so export never writes settings that a
  // re-import would immediately discard (a builder/hand-built doc could set either):
  // compat entries need a non-empty name, and the tab stop must be a positive number.
  const compatSettings = (opts.compatSettings ?? []).filter((c) => c.name.length > 0);
  const defaultTabStopPx =
    opts.defaultTabStopPx !== undefined && Number.isFinite(opts.defaultTabStopPx) && opts.defaultTabStopPx > 0
      ? opts.defaultTabStopPx
      : undefined;
  // w:displayBackgroundShape makes Word actually paint w:background (the page color).
  const compat =
    compatSettings.length > 0
      ? el(
          "w:compat",
          undefined,
          compatSettings
            .map((c) => el("w:compatSetting", { "w:name": c.name, "w:uri": c.uri, "w:val": c.val }))
            .join(""),
        )
      : "";
  const body =
    (opts.displayBackgroundShape ? el("w:displayBackgroundShape") : "") +
    (opts.evenAndOdd ? el("w:evenAndOddHeadersAndFooters") : "") +
    (defaultTabStopPx !== undefined ? el("w:defaultTabStop", { "w:val": Math.round(pxToTwips(defaultTabStopPx)) }) : "") +
    compat;
  return XML_DECL + el("w:settings", WML_NS, body);
}
