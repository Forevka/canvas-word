// Stage orchestration: bytes → Archive → styles/theme → document.xml (+ header/
// footer parts) → IR → Document. Pure (no worker/DOM references) so the whole
// pipeline runs under vitest in Node; worker.ts is just transport around this.

import type { Block, Paragraph, SectionProps } from "../../model/document";
import { findMainDocumentPart } from "./contentTypes";
import { parseDocumentXml, parseFootnotesXml, parseHeaderFooterXml } from "./documentParser";
import { buildStylesheet, collectUsedStyleIds, createMapper, mapSdts, type LinkResolver, type Mapper } from "./mapToModel";
import { createMediaStore, type MediaStore } from "./media";
import { parseNumberingXml, EMPTY_NUMBERING } from "./numbering";
import { findByType, parseRelationships, relsPartFor, type Relationships } from "./relationships";
import { createStyleResolver, parseStylesXml, resolveTableStyle, EMPTY_STYLES } from "./styles";
import { parseThemeXml, EMPTY_THEME } from "./theme";
import { openArchive, type Archive } from "./zip";
import { ImportError, WarningSink, type ImportPhase, type ImportResult, type IRSdtProps } from "./types";

export function runImport(
  bytes: Uint8Array,
  onProgress?: (phase: ImportPhase, pct: number) => void,
): ImportResult {
  const progress = (phase: ImportPhase, pct: number): void => onProgress?.(phase, pct);
  const warnings = new WarningSink();

  progress("unzip", 0);
  const archive = openArchive(bytes);
  progress("unzip", 1);

  const contentTypes = archive.text("[Content_Types].xml");
  const mainPart =
    (contentTypes !== undefined ? findMainDocumentPart(contentTypes) : undefined) ?? "word/document.xml";
  const documentXml = archive.text(mainPart);
  if (documentXml === undefined) {
    throw new ImportError("NO_DOCUMENT_PART", `Main document part "${mainPart}" not found in the archive.`);
  }

  // Relationships locate styles/theme/header/footer/media parts; conventional
  // names are the fallback for minimal producers that skip the rels part.
  const rels = relsOf(archive, mainPart);

  progress("styles", 0);
  const stylesXml = partByRelType(archive, rels, "styles") ?? archive.text("word/styles.xml");
  const styles = stylesXml !== undefined ? parseStylesXml(stylesXml, warnings) : EMPTY_STYLES;
  const themeXml = partByRelType(archive, rels, "theme") ?? archive.text("word/theme/theme1.xml");
  const theme = themeXml !== undefined ? parseThemeXml(themeXml) : EMPTY_THEME;
  const resolver = createStyleResolver(styles, theme);
  const numberingXml = partByRelType(archive, rels, "numbering") ?? archive.text("word/numbering.xml");
  const numbering = numberingXml !== undefined ? parseNumberingXml(numberingXml) : EMPTY_NUMBERING;
  progress("styles", 1);

  progress("parse", 0);
  const ir = parseDocumentXml(documentXml, mainPart, warnings);
  progress("parse", 1);

  progress("map", 0);
  // Reference page size for section-break geometry comparison: the document's
  // own section, falling back to Letter (so an explicit-size mid-doc section
  // still compares against the implied default).
  const refPgSize = {
    w: ir.section?.pageWidthTwips ?? 12240,
    h: ir.section?.pageHeightTwips ?? 15840,
  };
  const mapper = createMapper(
    warnings,
    resolver,
    ir.sdts,
    numbering,
    (styleId) => resolveTableStyle(styles, styleId),
    refPgSize,
  );
  const mediaStores: MediaStore[] = [];
  const mediaFor = (partRels: Relationships): MediaStore => {
    const store = createMediaStore(archive, partRels, warnings);
    mediaStores.push(store);
    return store;
  };

  const blocks = mapper.mapBlocks(ir.blocks, mediaFor(rels), linkResolverFor(rels));
  if (blocks.length === 0) blocks.push(mapper.emptyParagraph()); // caret needs a home
  const section = mapper.mapSection(ir.section);

  // Header/footer parts are full block stories with their OWN rels (images and
  // hyperlinks in a header resolve through header1.xml.rels, not the document's).
  // Their content controls join the document's sdt registry.
  const header = mapStory(archive, rels, ir.section?.headerRelId, mapper, mediaFor, warnings, ir.sdts);
  if (header) section.header = header;
  const footer = mapStory(archive, rels, ir.section?.footerRelId, mapper, mediaFor, warnings, ir.sdts);
  if (footer) section.footer = footer;

  // Footnotes: map the bodies of the notes actually referenced (in document
  // order, so numbering matches). Each note's own rels resolve its media/links.
  const footnotePart = partNameByRelType(rels, "footnotes") ?? "word/footnotes.xml";
  const footnoteXml = archive.text(footnotePart);
  const footnoteIR = footnoteXml !== undefined ? parseFootnotesXml(footnoteXml, footnotePart, warnings, ir.sdts) : new Map();
  const footnotes: Record<string, Paragraph[]> = {};
  if (footnoteIR.size > 0) {
    const footRels = relsOf(archive, footnotePart);
    const footMedia = mediaFor(footRels);
    const footLink = linkResolverFor(footRels);
    for (const { docxId, noteId } of mapper.footnoteRefs()) {
      const bodyIR = footnoteIR.get(docxId);
      if (!bodyIR) {
        warnings.add("footnote-missing", "A footnote reference had no matching note body.");
        continue;
      }
      const noteBlocks = mapper.mapBlocks(bodyIR, footMedia, footLink);
      const paras = noteBlocks.filter((b): b is Paragraph => b.kind === "paragraph");
      if (paras.length < noteBlocks.length) {
        warnings.add("footnote-tables", "Tables inside footnotes were dropped (footnotes hold paragraphs only).");
      }
      if (paras.length > 0) footnotes[noteId] = paras;
    }
  }
  progress("map", 1);

  const doc: ImportResult["doc"] = { section, blocks };
  if (Object.keys(footnotes).length > 0) doc.footnotes = footnotes;
  const sdts = mapSdts(ir.sdts);
  if (Object.keys(sdts).length > 0) doc.sdts = sdts;
  const lists = mapper.lists();
  if (Object.keys(lists).length > 0) doc.lists = lists;
  const bookmarks = mapper.bookmarks();
  if (Object.keys(bookmarks).length > 0) doc.bookmarks = bookmarks;
  // Style gallery: used paragraph styles (+ basedOn closure) with w:name labels.
  const stylesheet = buildStylesheet(styles, collectUsedStyleIds(ir.blocks));
  if (stylesheet) doc.stylesheet = stylesheet;

  return {
    doc,
    warnings: warnings.list,
    mediaUrls: mediaStores.flatMap((s) => s.urls()),
  };
}

function mapStory(
  archive: Archive,
  documentRels: Relationships,
  relId: string | undefined,
  mapper: Mapper,
  mediaFor: (rels: Relationships) => MediaStore,
  warnings: WarningSink,
  sdts: Record<string, IRSdtProps>,
): Block[] | undefined {
  if (!relId) return undefined;
  const rel = documentRels.get(relId);
  const xml = rel && !rel.external ? archive.text(rel.target) : undefined;
  if (rel === undefined || xml === undefined) {
    warnings.add("header-missing", "A referenced header/footer part was missing from the archive.");
    return undefined;
  }
  const partRels = relsOf(archive, rel.target);
  const ir = parseHeaderFooterXml(xml, rel.target, warnings, sdts);
  const blocks = mapper.mapBlocks(ir, mediaFor(partRels), linkResolverFor(partRels));
  return blocks.length > 0 ? blocks : undefined;
}

function relsOf(archive: Archive, partName: string): Relationships {
  const relsXml = archive.text(relsPartFor(partName));
  return relsXml !== undefined ? parseRelationships(relsXml, partName) : new Map();
}

/** Hyperlink r:id → URL. External rels carry the raw URL as target. */
function linkResolverFor(rels: Relationships): LinkResolver {
  return (relId) => {
    const rel = rels.get(relId);
    return rel?.external ? rel.target : undefined;
  };
}

function partByRelType(archive: Archive, rels: Relationships, kind: string): string | undefined {
  const rel = findByType(rels, kind);
  return rel && !rel.external ? archive.text(rel.target) : undefined;
}

/** Part NAME for a relationship type (footnotes need their rels resolved separately). */
function partNameByRelType(rels: Relationships, kind: string): string | undefined {
  const rel = findByType(rels, kind);
  return rel && !rel.external ? rel.target : undefined;
}
