// Document model -> .docx (OOXML zip). The inverse of the importer; pure and
// DOM-free, so it runs in the worker and on Node. Image bytes are supplied by the
// caller (blob: URLs resolve only on the main thread).

import { strToU8, zipSync, type Zippable } from "fflate";
import type { Block, Document } from "@cw/shared";
import type { ExportResult, ImageBytes } from "../types";
import { WarningSink } from "../warnings";
import { buildDocumentXml, type AddBandPart, type PartCtx } from "./documentXml";
import { contentTypesXml, CT } from "./contentTypes";
import { footnotesXml } from "./footnotesXml";
import { headerFooterXml } from "./headerFooterXml";
import { MediaManager } from "./mediaPack";
import { numberingXml } from "./numberingXml";
import { REL, RelManager, relsXml } from "./relationships";
import { settingsXml } from "./settingsXml";
import { stylesXml } from "./stylesXml";
import { XML_DECL, el } from "./xmlWrite";

export function writeDocx(doc: Document, images: ImageBytes = {}): ExportResult {
  const warnings = new WarningSink();
  const parts: Record<string, string | Uint8Array> = {};
  const overrides: [string, string][] = [["/word/document.xml", CT.document]];

  let idCounter = 1;
  const nextId = (): number => idCounter++;

  // Word-valid integer numIds for every list, shared between document & numbering.
  const listIdMap = new Map<string, number>();
  let nextNumId = 1;
  for (const listId of Object.keys(doc.lists ?? {})) listIdMap.set(listId, nextNumId++);

  // blockId -> bookmark names.
  const bookmarksByBlock = new Map<string, string[]>();
  for (const [name, blockId] of Object.entries(doc.bookmarks ?? {})) {
    const list = bookmarksByBlock.get(blockId) ?? [];
    list.push(name);
    bookmarksByBlock.set(blockId, list);
  }

  const sdts = doc.sdts ?? {};
  const bodyRels = new RelManager();
  const bodyMedia = new MediaManager(images);
  const bodyCtx: PartCtx = {
    rels: bodyRels,
    media: bodyMedia,
    warn: (c, d) => warnings.warn(c, d),
    sdts,
    nextId,
    bookmarksByBlock,
    listIdMap,
  };

  // Header/footer part factory — each band part gets its own rels + media.
  let bandCounter = 0;
  const addBand: AddBandPart = (blocks: Block[], kind: "header" | "footer"): string => {
    const idx = ++bandCounter;
    const file = `${kind}${idx}.xml`;
    const bandRels = new RelManager();
    const bandMedia = new MediaManager(images);
    const bandCtx: PartCtx = {
      rels: bandRels,
      media: bandMedia,
      warn: (c, d) => warnings.warn(c, d),
      sdts,
      nextId,
      bookmarksByBlock: new Map(),
      listIdMap,
    };
    parts[`word/${file}`] = headerFooterXml(blocks, kind, bandCtx);
    overrides.push([`/word/${file}`, kind === "header" ? CT.header : CT.footer]);
    if (!bandRels.isEmpty()) parts[`word/_rels/${file}.rels`] = relsXml(bandRels.list());
    for (const [path, bytes] of bandMedia.parts()) parts[`word/${path}`] = bytes;
    return bodyRels.add(kind === "header" ? REL.header : REL.footer, file);
  };

  // Body — also resolves mid-document section bands (writes their parts too).
  parts["word/document.xml"] = buildDocumentXml(doc.blocks, doc.section, bodyCtx, addBand);

  // styles.xml
  if (doc.stylesheet && doc.stylesheet.styles.length > 0) {
    parts["word/styles.xml"] = stylesXml(doc.stylesheet);
    overrides.push(["/word/styles.xml", CT.styles]);
    bodyRels.add(REL.styles, "styles.xml");
  }

  // numbering.xml
  if (doc.lists && Object.keys(doc.lists).length > 0) {
    parts["word/numbering.xml"] = numberingXml(doc.lists, listIdMap);
    overrides.push(["/word/numbering.xml", CT.numbering]);
    bodyRels.add(REL.numbering, "numbering.xml");
  }

  // footnotes.xml
  if (doc.footnotes && Object.keys(doc.footnotes).length > 0) {
    parts["word/footnotes.xml"] = footnotesXml(doc.footnotes, bodyCtx);
    overrides.push(["/word/footnotes.xml", CT.footnotes]);
    bodyRels.add(REL.footnotes, "footnotes.xml");
  }

  // settings.xml (always — carries the even/odd flag).
  const evenAndOdd = sectionsHaveEvenBands(doc);
  parts["word/settings.xml"] = settingsXml(evenAndOdd);
  overrides.push(["/word/settings.xml", CT.settings]);
  bodyRels.add(REL.settings, "settings.xml");

  // Body media + rels (after all rels are registered).
  for (const [path, bytes] of bodyMedia.parts()) parts[`word/${path}`] = bytes;
  parts["word/_rels/document.xml.rels"] = relsXml(bodyRels.list());

  // Package-level parts.
  parts["[Content_Types].xml"] = contentTypesXml(overrides, bodyMedia.extensions());
  parts["_rels/.rels"] =
    XML_DECL +
    el(
      "Relationships",
      { xmlns: "http://schemas.openxmlformats.org/package/2006/relationships" },
      el("Relationship", {
        Id: "rId1",
        Type: REL.officeDocument,
        Target: "word/document.xml",
      }),
    );

  const zippable: Zippable = {};
  for (const [name, content] of Object.entries(parts)) {
    zippable[name] = typeof content === "string" ? strToU8(content) : content;
  }
  const bytes = zipSync(zippable);
  return { bytes, warnings: warnings.list() };
}

function sectionsHaveEvenBands(doc: Document): boolean {
  const s = doc.section;
  if (s.headerEven || s.footerEven) return true;
  for (const b of doc.blocks) {
    if (b.kind === "paragraph" && b.style.sectionBreak) {
      const p = b.style.sectionBreak.props;
      if (p.headerEven || p.footerEven) return true;
    }
  }
  return false;
}
