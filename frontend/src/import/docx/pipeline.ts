// Stage orchestration: bytes → Archive → styles/theme → document.xml (+ header/
// footer parts) → IR → Document. Pure (no worker/DOM references) so the whole
// pipeline runs under vitest in Node; worker.ts is just transport around this.

import type { Block, BandContainer, Paragraph, SectionProps } from "@cw/shared";
import { markImportedTocEntries } from "@cw/shared";
import { findMainDocumentPart } from "./contentTypes";
import { parseDocumentXml, parseFootnotesXml, parseHeaderFooterXml } from "./documentParser";
import { buildStylesheet, buildTableStyles, createMapper, mapSdts, type LinkResolver, type Mapper } from "./mapToModel";
import { createMediaStore, type MediaStore } from "./media";
import { parseNumberingXml, EMPTY_NUMBERING } from "./numbering";
import { findByType, parseRelationships, relsPartFor, type Relationships } from "./relationships";
import { createStyleResolver, parseStylesXml, resolveTableStyle, EMPTY_STYLES } from "./styles";
import { parseThemeXml, EMPTY_THEME } from "./theme";
import { attr, el, els, numAttr, parseXml, rootEl } from "./xml";
import { round2, twipsToPx } from "./units";
import { openArchive, type Archive } from "./zip";
import { ImportError, WarningSink, type BandRefs, type ImportMedia, type ImportPhase, type ImportResult, type IRSdtProps, type MediaCollector, type RunImportOpts } from "./types";

export function runImport(
  bytes: Uint8Array,
  onProgress?: (phase: ImportPhase, pct: number) => void,
  opts?: RunImportOpts,
): ImportResult {
  const progress = (phase: ImportPhase, pct: number): void =>
    (opts?.onProgress ?? onProgress)?.(phase, pct);
  const warnings = new WarningSink();

  // Node/backend path: collect raw image bytes instead of minting blob: URLs.
  // Each call returns a synthetic src ("cw-media:N") the mapper stamps on the
  // ImageBlock, so the caller can match bytes → block and swap in a mediaId.
  const mediaRecords: ImportMedia[] = [];
  const collector: MediaCollector | undefined = opts?.collectMediaBytes
    ? {
        add(b, mime) {
          const src = `cw-media:${mediaRecords.length}`;
          mediaRecords.push({ src, bytes: b, mime });
          return src;
        },
      }
    : undefined;

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
    const store = createMediaStore(archive, partRels, warnings, collector);
    mediaStores.push(store);
    return store;
  };

  const blocks = mapper.mapBlocks(ir.blocks, mediaFor(rels), linkResolverFor(rels));
  if (blocks.length === 0) blocks.push(mapper.emptyParagraph()); // caret needs a home
  const section = mapper.mapSection(ir.section);
  // w:background is document-global → the body section's page fill.
  if (ir.pageColorHex) section.pageColorHex = `#${ir.pageColorHex}`;

  // Header/footer parts are full block stories with their OWN rels (images and
  // hyperlinks in a header resolve through header1.xml.rels, not the document's).
  // Their content controls join the document's sdt registry. The first/even
  // variants only apply when Word would use them (w:titlePg / settings evenAndOdd).
  const settingsXml = partByRelType(archive, rels, "settings") ?? archive.text("word/settings.xml");
  const settings = settingsXml !== undefined ? parseSettings(settingsXml) : { evenAndOdd: false };
  const evenAndOdd = settings.evenAndOdd;
  mapBands(section, "header", ir.section?.headerRefs, ir.section?.titlePg ?? false, evenAndOdd, archive, rels, mapper, mediaFor, warnings, ir.sdts);
  mapBands(section, "footer", ir.section?.footerRefs, ir.section?.titlePg ?? false, evenAndOdd, archive, rels, mapper, mediaFor, warnings, ir.sdts);

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
  // Style gallery: ALL defined paragraph/character styles (w:name labels) — so
  // authored-but-unapplied styles survive a round-trip (matches Word).
  const stylesheet = buildStylesheet(styles);
  if (stylesheet) doc.stylesheet = stylesheet;
  // All defined table styles → Document.tableStyles.
  const tableStyles = buildTableStyles(styles);
  if (tableStyles) doc.tableStyles = tableStyles;

  // TOC field: capture its instruction (for verbatim re-emit) and mark the
  // flattened entry paragraphs as live tocEntry (stripping their cached numbers)
  // so export round-trips them as a real, F9-refreshable field, not frozen text.
  // When the field has NO entries (e.g. a C# pipeline emitting an empty TOC field),
  // record the field's block (captured at map time, not by ordinal) so a headless
  // render can BUILD the entries in place.
  const toc = mapper.tocField();
  if (toc) {
    doc.tocInstruction = toc.instruction;
    if (markImportedTocEntries(doc) === 0) doc.tocAnchorBlockId = toc.blockId;
  }

  // Custom (non-built-in) fields captured during the body walk — their result
  // blocks already carry the matching fieldId (see mapToModel).
  if (ir.fields) doc.fields = ir.fields;

  // Document-level settings.xml bits: a non-default tab interval honored at layout,
  // and compat triples round-tripped verbatim.
  if (settings.defaultTabStopPx !== undefined) doc.defaultTabStopPx = settings.defaultTabStopPx;
  if (settings.compatSettings) doc.compatSettings = settings.compatSettings;

  return {
    doc,
    warnings: warnings.list,
    mediaUrls: mediaStores.flatMap((s) => s.urls()),
    media: mediaRecords,
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

/** Map a band's default/first/even references onto the section's slots. First
 *  applies only with w:titlePg; even only with settings evenAndOdd — matching
 *  when Word actually renders them. */
function mapBands(
  section: SectionProps,
  kind: "header" | "footer",
  refs: BandRefs | undefined,
  titlePg: boolean,
  evenAndOdd: boolean,
  archive: Archive,
  rels: Relationships,
  mapper: Mapper,
  mediaFor: (rels: Relationships) => MediaStore,
  warnings: WarningSink,
  sdts: Record<string, IRSdtProps>,
): void {
  if (!refs) return;
  const slot = (variant: "" | "First" | "Even"): BandContainer => `${kind}${variant}` as BandContainer;
  const set = (container: BandContainer, relId: string | undefined): void => {
    const story = mapStory(archive, rels, relId, mapper, mediaFor, warnings, sdts);
    if (story) section[container] = story;
  };
  set(slot(""), refs.default);
  if (titlePg) set(slot("First"), refs.first);
  if (evenAndOdd) set(slot("Even"), refs.even);
}

interface ParsedSettings {
  /** w:evenAndOddHeadersAndFooters — even-page bands only apply when on. */
  evenAndOdd: boolean;
  /** w:defaultTabStop (twips) → px — the layout's tab interval past the last stop. */
  defaultTabStopPx?: number;
  /** w:compat/w:compatSetting triples, round-tripped verbatim. */
  compatSettings?: { name: string; uri: string; val: string }[];
}

/** Decode the document-level settings.xml bits the model cares about. */
function parseSettings(settingsXml: string): ParsedSettings {
  const root = rootEl(parseXml(settingsXml, "settings.xml"), "w:settings");
  if (!root) return { evenAndOdd: false };
  const out: ParsedSettings = { evenAndOdd: !!el(root, "w:evenAndOddHeadersAndFooters") };
  const twips = numAttr(el(root, "w:defaultTabStop"), "w:val");
  // Word's own default is 720 twips; only carry an explicit, non-default value so a
  // plain document keeps the engine's fallback constant (no needless model field).
  if (twips !== undefined && twips > 0 && twips !== 720) out.defaultTabStopPx = round2(twipsToPx(twips));
  const compat = el(root, "w:compat");
  if (compat) {
    const settings = els(compat, "w:compatSetting")
      .map((c) => ({ name: attr(c, "w:name") ?? "", uri: attr(c, "w:uri") ?? "", val: attr(c, "w:val") ?? "" }))
      .filter((c) => c.name !== "");
    if (settings.length > 0) out.compatSettings = settings;
  }
  return out;
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
