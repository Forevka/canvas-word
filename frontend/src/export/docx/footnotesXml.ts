// footnotes.xml — the standard separator pseudo-notes plus one w:footnote per
// real note. Keys in Document.footnotes are "fn<docxId>"; the id Word stores is
// <docxId>, matching the w:footnoteReference written into the document.

import type { Paragraph } from "@cw/shared";
import { blockXml, type PartCtx } from "./documentXml";
import { el, WML_NS, XML_DECL } from "./xmlWrite";

export function footnotesXml(footnotes: Record<string, Paragraph[]>, ctx: PartCtx): string {
  const sep =
    el("w:footnote", { "w:type": "separator", "w:id": -1 }, el("w:p", undefined, el("w:r", undefined, el("w:separator")))) +
    el("w:footnote", { "w:type": "continuationSeparator", "w:id": 0 }, el("w:p", undefined, el("w:r", undefined, el("w:continuationSeparator"))));
  const notes = Object.entries(footnotes)
    .map(([noteId, paras]) => {
      const docxId = noteId.replace(/^fn/, "");
      const body = paras.length > 0 ? paras.map((p) => blockXml(p, ctx)).join("") : el("w:p");
      return el("w:footnote", { "w:id": docxId }, body);
    })
    .join("");
  return XML_DECL + el("w:footnotes", WML_NS, sep + notes);
}
