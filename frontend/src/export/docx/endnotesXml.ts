// endnotes.xml — the standard separator pseudo-notes plus one w:endnote per
// real note. Keys in Document.endnotes are "en<docxId>"; the id Word stores is
// <docxId>, matching the w:endnoteReference written into the document. Mirrors
// footnotesXml.ts.

import type { Paragraph } from "@cw/shared";
import { emitBlocks, type PartCtx } from "./documentXml";
import { el, WML_NS, XML_DECL } from "./xmlWrite";

export function endnotesXml(endnotes: Record<string, Paragraph[]>, ctx: PartCtx): string {
  const sep =
    el("w:endnote", { "w:type": "separator", "w:id": -1 }, el("w:p", undefined, el("w:r", undefined, el("w:separator")))) +
    el("w:endnote", { "w:type": "continuationSeparator", "w:id": 0 }, el("w:p", undefined, el("w:r", undefined, el("w:continuationSeparator"))));
  const notes = Object.entries(endnotes)
    .map(([noteId, paras]) => {
      const docxId = noteId.replace(/^en/, "");
      const body = paras.length > 0 ? emitBlocks(paras, 0, ctx) : el("w:p");
      return el("w:endnote", { "w:id": docxId }, body);
    })
    .join("");
  return XML_DECL + el("w:endnotes", WML_NS, sep + notes);
}
