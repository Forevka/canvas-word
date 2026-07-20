// OOXML enforces a strict xsd:sequence on the children of every property element
// (w:pPr, w:rPr, w:tcPr, w:sectPr, …). The importer is order-independent and lenient
// consumers (Google Docs, LibreOffice) don't care, but Word validates against the
// schema and rejects the whole document on the FIRST out-of-order child ("Word
// experienced an error trying to open the file"). See issue #180.
//
// Rather than hand-order every push in the serializers (fragile — a future field
// added in the wrong spot silently re-breaks Word), each serializer collects its
// child elements in any convenient order and runs the list through `orderChildren`
// with the canonical order for its container. One authoritative table per type.

/** Leading element name (`w:spacing`, `w15:docId`) of a serialized child. Each child
 *  passed to `orderChildren` must be exactly ONE top-level element, so this reads the
 *  opening tag only — nested descendants are ignored. */
function tagOf(elementXml: string): string {
  const m = /^<([\w]+:[\w]+)/.exec(elementXml);
  return m ? m[1]! : "";
}

/** Stably reorder serialized child elements to the schema sequence in `order`.
 *  Elements sharing a tag keep their insertion order (e.g. the default/first/even
 *  header references); an unrecognized tag sorts to the end, preserving order. */
export function orderChildren(children: string[], order: readonly string[]): string[] {
  const rank = new Map(order.map((t, i) => [t, i]));
  return children
    .map((c, i) => ({ c, i, r: rank.get(tagOf(c)) ?? Number.MAX_SAFE_INTEGER }))
    .sort((a, b) => a.r - b.r || a.i - b.i)
    .map((x) => x.c);
}

// --- Canonical child sequences (ECMA-376 / ISO-29500). Only the elements the
// exporter can emit need to be listed, but the full sequence is kept so the tables
// double as documentation and tolerate future additions. ---

/** CT_PPr — paragraph properties (w:pPr). */
export const PPR_ORDER = [
  "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr", "w:widowControl",
  "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens",
  "w:kinsoku", "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
  "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing",
  "w:mirrorIndents", "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
  "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
] as const;

/** CT_RPr / EG_RPrBase — run properties (w:rPr). */
export const RPR_ORDER = [
  "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps", "w:strike",
  "w:dstrike", "w:outline", "w:shadow", "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid",
  "w:vanish", "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position", "w:sz", "w:szCs",
  "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs",
  "w:em", "w:lang", "w:eastAsianLayout", "w:specVanish", "w:oMath", "w:rPrChange",
] as const;

/** CT_TcPr — table-cell properties (w:tcPr). */
export const TCPR_ORDER = [
  "w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge", "w:tcBorders", "w:shd", "w:noWrap",
  "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark", "w:cellIns", "w:cellDel",
  "w:cellMerge", "w:tcPrChange",
] as const;

/** CT_SectPr — section properties (w:sectPr). The header/footer references form an
 *  unbounded leading choice, so they sort ahead of everything else. */
export const SECTPR_ORDER = [
  "w:headerReference", "w:footerReference", "w:footnotePr", "w:endnotePr", "w:type", "w:pgSz",
  "w:pgMar", "w:paperSrc", "w:pgBorders", "w:lnNumType", "w:pgNumType", "w:cols", "w:formProt",
  "w:vAlign", "w:noEndnote", "w:titlePg", "w:textDirection", "w:bidi", "w:rtlGutter", "w:docGrid",
  "w:printerSettings", "w:sectPrChange",
] as const;
