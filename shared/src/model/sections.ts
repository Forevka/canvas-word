// Section resolution — pure model logic (no layout, no DOM). A section-break
// paragraph TERMINATES a section (OOXML sectPr placement); blocks after the
// last break belong to `doc.section`. Absent patch fields inherit from
// `doc.section` ("link to previous"). Lives in the shared core so both the
// layout engine and the document query API resolve sections the same way.

import type { Document, SectionBreakType, SectionPatch, SectionProps } from "./document";
import { BAND_CONTAINERS } from "./document";

export function effectiveSection(base: SectionProps, patch: SectionPatch): SectionProps {
  const out: SectionProps = {
    pageWidthPx: patch.pageWidthPx ?? base.pageWidthPx,
    pageHeightPx: patch.pageHeightPx ?? base.pageHeightPx,
    marginPx: patch.marginPx ?? base.marginPx,
  };
  // columns: undefined = inherit, null = explicitly single-column
  const columns = patch.columns === undefined ? base.columns : (patch.columns ?? undefined);
  if (columns) out.columns = columns;
  // page-number restart is a section's OWN property — never inherited
  if (patch.pageNumberStart !== undefined) out.pageNumberStart = patch.pageNumberStart;
  // Band distances inherit from the document section (the report sets them once).
  const headerDist = patch.headerDistancePx ?? base.headerDistancePx;
  if (headerDist !== undefined) out.headerDistancePx = headerDist;
  const footerDist = patch.footerDistancePx ?? base.footerDistancePx;
  if (footerDist !== undefined) out.footerDistancePx = footerDist;
  // Page fill & borders inherit from the document section unless overridden.
  const pageColorHex = patch.pageColorHex ?? base.pageColorHex;
  if (pageColorHex !== undefined) out.pageColorHex = pageColorHex;
  const pageBorders = patch.pageBorders ?? base.pageBorders;
  if (pageBorders !== undefined) out.pageBorders = pageBorders;
  // Line numbering is a section's OWN property (like page-number restart): a
  // section either declares its own w:lnNumType or has none — it never inherits.
  if (patch.lineNumbering !== undefined) out.lineNumbering = patch.lineNumbering;
  for (const key of BAND_CONTAINERS) {
    const blocks = patch[key] ?? base[key];
    if (blocks) out[key] = blocks;
  }
  return out;
}

export interface ResolvedSection {
  /** 0-based position of this section within the document's section list. */
  index: number;
  props: SectionProps;
  /** Index (inclusive) of this section's FIRST top-level block. */
  startBlock: number;
  /** Index (inclusive) of this section's LAST top-level block. */
  endBlock: number;
  /** The OOXML w:type that governs how THIS section's first page begins
   *  ("nextPage"/"evenPage"/"oddPage"). Read when the engine starts the section. */
  breakType: SectionBreakType;
}

export function resolveSections(doc: Document): ResolvedSection[] {
  const out: ResolvedSection[] = [];
  let startBlock = 0;
  for (let i = 0; i < doc.blocks.length; i++) {
    const b = doc.blocks[i]!;
    if (b.kind === "paragraph" && b.style.sectionBreak) {
      out.push({
        index: out.length,
        props: effectiveSection(doc.section, b.style.sectionBreak.props),
        startBlock,
        endBlock: i,
        breakType: b.style.sectionBreak.type,
      });
      startBlock = i + 1;
    }
  }
  // The trailing body section keeps its own start type (even/odd parity), so a
  // document whose final section begins on a parity page is honored, not flattened.
  out.push({
    index: out.length,
    props: doc.section,
    startBlock,
    endBlock: doc.blocks.length - 1,
    breakType: doc.section.breakType ?? "nextPage",
  });
  return out;
}
