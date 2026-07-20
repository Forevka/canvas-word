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
  /** w15:docId GUID to re-emit verbatim (declares the w15 namespace). */
  docId?: string;
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
  const docId = opts.docId && opts.docId.length > 0 ? opts.docId : undefined;
  // CT_Settings is an xsd:sequence; Word rejects out-of-order children. The schema
  // order of what we emit is displayBackgroundShape(7) → defaultTabStop(39) →
  // evenAndOddHeaders(48) → compat(81) → w15:docId (issue #180).
  const body =
    (opts.displayBackgroundShape ? el("w:displayBackgroundShape") : "") +
    (defaultTabStopPx !== undefined ? el("w:defaultTabStop", { "w:val": Math.round(pxToTwips(defaultTabStopPx)) }) : "") +
    (opts.evenAndOdd ? el("w:evenAndOddHeadersAndFooters") : "") +
    compat +
    // w15:docId (Office 2013+) rides at the end of the settings sequence via the w15
    // extension; the namespace + mc:Ignorable are declared on the root below.
    (docId !== undefined ? el("w15:docId", { "w15:val": docId }) : "");
  const ns = {
    ...WML_NS,
    "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "mc:Ignorable": "w14 w15",
  };
  return XML_DECL + el("w:settings", ns, body);
}
