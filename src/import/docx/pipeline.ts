// Stage orchestration: bytes → Archive → styles/theme → document.xml (+ header/
// footer parts) → IR → Document. Pure (no worker/DOM references) so the whole
// pipeline runs under vitest in Node; worker.ts is just transport around this.

import type { Block, SectionProps } from "../../model/document";
import { findMainDocumentPart } from "./contentTypes";
import { parseDocumentXml, parseHeaderFooterXml } from "./documentParser";
import { createMapper, type Mapper } from "./mapToModel";
import { createMediaStore, type MediaStore } from "./media";
import { findByType, parseRelationships, relsPartFor, type Relationships } from "./relationships";
import { createStyleResolver, parseStylesXml, EMPTY_STYLES } from "./styles";
import { parseThemeXml, EMPTY_THEME } from "./theme";
import { openArchive, type Archive } from "./zip";
import { ImportError, WarningSink, type ImportPhase, type ImportResult } from "./types";

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
  progress("styles", 1);

  progress("parse", 0);
  const ir = parseDocumentXml(documentXml, mainPart, warnings);
  progress("parse", 1);

  progress("map", 0);
  const mapper = createMapper(warnings, resolver);
  const mediaStores: MediaStore[] = [];
  const mediaFor = (partRels: Relationships): MediaStore => {
    const store = createMediaStore(archive, partRels, warnings);
    mediaStores.push(store);
    return store;
  };

  const blocks = mapper.mapBlocks(ir.blocks, mediaFor(rels));
  if (blocks.length === 0) blocks.push(mapper.emptyParagraph()); // caret needs a home
  const section = mapper.mapSection(ir.section);

  // Header/footer parts are full block stories with their OWN rels (images in
  // a header resolve through header1.xml.rels, not the document's).
  const header = mapStory(archive, rels, ir.section?.headerRelId, mapper, mediaFor, warnings);
  if (header) section.header = header;
  const footer = mapStory(archive, rels, ir.section?.footerRelId, mapper, mediaFor, warnings);
  if (footer) section.footer = footer;
  progress("map", 1);

  return {
    doc: { section, blocks },
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
): Block[] | undefined {
  if (!relId) return undefined;
  const rel = documentRels.get(relId);
  const xml = rel && !rel.external ? archive.text(rel.target) : undefined;
  if (rel === undefined || xml === undefined) {
    warnings.add("header-missing", "A referenced header/footer part was missing from the archive.");
    return undefined;
  }
  const ir = parseHeaderFooterXml(xml, rel.target, warnings);
  const blocks = mapper.mapBlocks(ir, mediaFor(relsOf(archive, rel.target)));
  return blocks.length > 0 ? blocks : undefined;
}

function relsOf(archive: Archive, partName: string): Relationships {
  const relsXml = archive.text(relsPartFor(partName));
  return relsXml !== undefined ? parseRelationships(relsXml, partName) : new Map();
}

function partByRelType(archive: Archive, rels: Relationships, kind: string): string | undefined {
  const rel = findByType(rels, kind);
  return rel && !rel.external ? archive.text(rel.target) : undefined;
}
