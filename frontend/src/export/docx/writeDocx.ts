// Document model -> .docx (OOXML zip). The inverse of the importer; pure and
// DOM-free, so it runs in the worker and on Node. Image bytes are supplied by the
// caller (blob: URLs resolve only on the main thread).

import { strToU8, zipSync, type Zippable } from "fflate";
import type { Block, Document } from "@cw/shared";
import type { ExportResult, ImageBytes } from "../types";
import { WarningSink } from "../warnings";
import { buildDocumentXml, type AddBandPart, type ExportBookmarkMark, type PartCtx } from "./documentXml";
import { contentTypesXml, CT } from "./contentTypes";
import { footnotesXml } from "./footnotesXml";
import { headerFooterXml } from "./headerFooterXml";
import { MediaManager } from "./mediaPack";
import { numberingXml } from "./numberingXml";
import { REL, RelManager, relsXml } from "./relationships";
import { settingsXml } from "./settingsXml";
import { stylesXml } from "./stylesXml";
import { XML_DECL, el } from "./xmlWrite";

export function writeDocx(
  doc: Document,
  images: ImageBytes = {},
  tocPages?: Map<string, number>,
): ExportResult {
  const warnings = new WarningSink();
  const parts: Record<string, string | Uint8Array> = {};
  const overrides: [string, string][] = [["/word/document.xml", CT.document]];

  let idCounter = 1;
  const nextId = (): number => idCounter++;

  // Word-valid integer numIds for every list, shared between document & numbering.
  const listIdMap = new Map<string, number>();
  let nextNumId = 1;
  for (const listId of Object.keys(doc.lists ?? {})) listIdMap.set(listId, nextNumId++);

  // blockId -> bookmark start/end markers. Keyed by block id, so band (header/
  // footer) blocks resolve when the same map is handed to the band parts. A
  // shared integer id pairs each start with its end.
  const bookmarksByBlock = new Map<string, ExportBookmarkMark[]>();
  let bmId = 0;
  const pushMark = (blockId: string, mark: ExportBookmarkMark): void => {
    const list = bookmarksByBlock.get(blockId) ?? [];
    list.push(mark);
    bookmarksByBlock.set(blockId, list);
  };
  for (const [name, range] of Object.entries(doc.bookmarks ?? {})) {
    const id = bmId++;
    pushMark(range.start.blockId, { id, name, kind: "start", offset: range.start.offset });
    pushMark(range.end.blockId, { id, kind: "end", offset: range.end.offset });
  }

  // TOC export: each entry's PAGEREF needs a bookmark on its target heading. Reuse
  // an existing one (imported docs have _Toc bookmarks) or synthesize a point
  // bookmark — but only for headings that actually exist (dangling entries skip).
  const paraIds = new Set<string>();
  const collectIds = (blocks: Block[]): void => {
    for (const b of blocks) {
      if (b.kind === "paragraph") paraIds.add(b.id);
      else if (b.kind === "table") for (const r of b.rows) for (const c of r.cells) collectIds(c.blocks);
    }
  };
  collectIds(doc.blocks);
  const bmNameByBlock = new Map<string, string>();
  for (const [name, range] of Object.entries(doc.bookmarks ?? {})) {
    if (!bmNameByBlock.has(range.start.blockId)) bmNameByBlock.set(range.start.blockId, name);
  }
  let synthBm = 0;
  const ensureTocBookmark = (blockId: string): string | undefined => {
    if (!paraIds.has(blockId)) return undefined; // dangling target
    const existing = bmNameByBlock.get(blockId);
    if (existing) return existing;
    const name = `_Toc9${String(synthBm++).padStart(8, "0")}`; // avoid clashing with real _TocNNN
    bmNameByBlock.set(blockId, name);
    const id = bmId++;
    pushMark(blockId, { id, name, kind: "start", offset: 0 });
    pushMark(blockId, { id, kind: "end", offset: 0 });
    return name;
  };

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
    ...(tocPages ? { tocPages } : {}),
    ...(doc.tocInstruction ? { tocInstruction: doc.tocInstruction } : {}),
    ...(doc.fields ? { fields: doc.fields } : {}),
    ensureTocBookmark,
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
      bookmarksByBlock, // same map — band block ids resolve here, so band bookmarks export
      listIdMap,
      fieldTokens: true, // {page}/{pages} -> live PAGE/NUMPAGES fields in bands
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
