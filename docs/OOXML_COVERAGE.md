# OOXML coverage

A living map of which **WordprocessingML** (`.docx`) elements and attributes the
import → model → export pipeline round-trips. It is the companion to the lossy-mapping
table in [IMPORT.md](IMPORT.md): that table explains _policy_ for the hard gaps; this
tree is the _element-by-element_ inventory.

**Keep it current.** Whenever you add support for an element, change how one is decoded/
emitted, or deliberately drop one, update the matching line here in the same PR. New
round-trip work (e.g. #62, #147, #161, #167) should leave this file reflecting reality.

## Legend

| Mark | Meaning |
|---|---|
| ✅ | **Covered** — decoded on import and re-emitted on export (round-trips) |
| 🔷 | **Round-trip only** — preserved byte-for-byte but has no layout/paint effect in our model |
| ⚠️ | **Partial** — some attributes/children covered, others dropped (see note) |
| ❌ | **Dropped** — silently lost on import (candidate for a future fidelity PR) |
| 🚫 | **Out of scope** — intentionally not modeled (documented decision) |

"Round-trip only" (🔷) still counts as covered — it survives a save/reopen; it just
doesn't change layout (our model implements no behavior for it). ❌ is the actionable
category — the same class as issues #161/#167/#168.

Primary decoders: `frontend/src/import/docx/props.ts` (`decodeRunProps`, `decodeParaProps`),
`frontend/src/import/docx/documentParser.ts` (run/paragraph/table content),
`frontend/src/import/docx/mapToModel.ts` (IR → model). Exporters:
`frontend/src/export/docx/documentXml.ts` and `styleProps.ts`.

---

## `w:r` → `w:rPr` (run properties)

```
w:rPr
├─ w:rStyle ................................ ✅ CharStyle.charStyleId
├─ w:b (bold) ............................. ✅
├─ w:bCs (complex-script bold) ............ ❌
├─ w:i (italic) ........................... ✅
├─ w:iCs (complex-script italic) .......... ❌
├─ w:caps ................................. ✅
├─ w:smallCaps ............................ ✅
├─ w:strike ............................... ✅ CharStyle.strikethrough
├─ w:dstrike (double strike) .............. ✅ doubleStrikethrough
├─ w:outline .............................. 🔷 outline
├─ w:shadow ............................... 🔷 shadow
├─ w:emboss ............................... 🔷 emboss
├─ w:imprint .............................. 🔷 imprint
├─ w:noProof .............................. ❌
├─ w:snapToGrid ........................... 🔷 snapToGrid (inter-character; #161)
├─ w:vanish ............................... ✅ hidden (preserved, never painted)
├─ w:specVanish ........................... ❌
├─ w:color ................................ ✅ @w:val + @w:themeColor/@w:themeTint/@w:themeShade
├─ w:spacing (char tracking) .............. ✅ letterSpacingPx
├─ w:w (width scaling %) .................. ✅ widthScalePct
├─ w:kern (kerning threshold) ............. ✅ kerningMinPx
├─ w:position (baseline shift) ............ ✅ positionPx
├─ w:sz ................................... ✅ fontSizePx
├─ w:szCs (complex-script size) ........... ⚠️ export emits (= w:sz); import ignores a divergent value
├─ w:highlight ............................ ✅ incl. explicit "none" clear (#155)
├─ w:shd (character shading) .............. ❌  (paragraph & cell w:shd ARE covered)
├─ w:u (underline) ........................ ✅ @w:val style + @w:color + @w:themeColor
├─ w:em (emphasis mark) ................... ✅ emphasisMark
├─ w:bdr (run border) ..................... ✅ runBorder (reuses CellBorder)
├─ w:rFonts .............................. ✅ @w:ascii/@w:hAnsi/@w:cs/@w:eastAsia + *Theme
├─ w:vertAlign (super/sub) ................ ✅ verticalAlign
├─ w:rtl .................................. ✅ rtl (explicit w:rtl="0" kept)
├─ w:lang (@val/@eastAsia/@bidi) .......... ✅ CharStyle.lang (run + style + docDefaults; #168)
├─ w:effect (text animation) .............. ❌  (legacy blink/shimmer)
├─ w:fitText .............................. 🔷 fitTextPx
├─ w:eastAsianLayout (combine/vert) ....... ❌
└─ w:oMath (inline equation) .............. ✅ CharStyle.equation (MathML)
```

### `w:r` run content (decoded in `documentParser.ts`)

```
w:r
├─ w:t (text) ............................. ✅
├─ w:tab .................................. ✅ laid out at w:tabs stops
├─ w:br (@w:type) ......................... ✅ page/column → break; textWrapping → paragraph split
├─ w:cr ................................... ✅ (line break)
├─ w:sym (symbol glyph) ................... ✅ symbol {font,char}; font set so it paints
├─ w:drawing (inline/anchor image) ........ ✅ see Drawing section
├─ w:fldChar / w:instrText ................ ✅ complex fields — see Fields
├─ w:footnoteReference .................... ✅ footnoteRef
├─ w:endnoteReference ..................... ✅ endnoteRef
├─ w:noBreakHyphen ........................ ⚠️ treated as a normal hyphen
├─ w:softHyphen ........................... ⚠️ dropped (invisible)
├─ w:ptab (absolute-position tab) ......... ❌
├─ w:ruby (phonetic guide) ................ ❌
├─ w:object (OLE) ......................... 🚫 out of scope
└─ w:pict (VML shape) ..................... 🚫 out of scope
```

---

## `w:p` → `w:pPr` (paragraph properties)

```
w:pPr
├─ w:pStyle ............................... ✅ namedStyle (via style cascade)
├─ w:keepNext ............................. ✅ keepWithNext
├─ w:keepLines ............................ ✅ keepLinesTogether
├─ w:pageBreakBefore ...................... ✅ pageBreakBefore
├─ w:widowControl ......................... ✅ widowControl (honored by pagination)
├─ w:suppressLineNumbers .................. 🔷 suppressLineNumbers
├─ w:suppressAutoHyphens .................. 🔷 suppressAutoHyphens (#161)
├─ w:kinsoku .............................. 🔷 kinsoku (#161)
├─ w:overflowPunct ........................ 🔷 overflowPunct (#161)
├─ w:snapToGrid ........................... 🔷 snapToGrid (inter-line; #161)
├─ w:wordWrap ............................. 🔷 wordWrap (#167)
├─ w:topLinePunct ......................... 🔷 topLinePunct (#167)
├─ w:autoSpaceDE .......................... 🔷 autoSpaceDE (#167)
├─ w:autoSpaceDN .......................... 🔷 autoSpaceDN (#167)
├─ w:mirrorIndents ........................ 🔷 mirrorIndents
├─ w:adjustRightInd ....................... 🔷 adjustRightInd
├─ w:contextualSpacing .................... ✅ contextualSpacing
├─ w:numPr (w:numId + w:ilvl) ............. ✅ list; numId=0 = explicit clear (#152)
├─ w:pBdr (all edges + between) ........... ✅ borders; empty w:pBdr = explicit clear (#153)
├─ w:shd (paragraph shading) .............. ✅ shading; empty w:shd = explicit clear (#147)
├─ w:tabs ................................. ✅ pos/val/leader; "clear" kept (#154); "bar" skipped ⚠️
├─ w:spacing ............................. ✅ before/after/line/lineRule + before/afterAutospacing (#160)
├─ w:ind ................................. ✅ left/right/firstLine/hanging (+ start/end)
├─ w:jc (alignment) ....................... ✅ incl. start/end/both/distribute aliases
├─ w:textAlignment (vertical in line) ..... ✅ textAlignment
├─ w:outlineLvl ........................... ✅ outlineLevel (drives TOC/heading detection)
├─ w:bidi (RTL paragraph) ................. ✅ direction (explicit w:bidi="0" kept)
├─ w:sectPr (section on paragraph) ........ ✅ see Sections
├─ w:rPr (paragraph-mark run props) ....... ✅ markRunProps
├─ w:framePr (text frame) ................. ❌
├─ w:textboxTightWrap ..................... ❌
├─ w:cnfStyle (conditional table fmt) ..... ❌
├─ w:divId (HTML div) ..................... ❌
└─ w:suppressOverlap ...................... ❌
```

`w:p` attributes: `@w14:paraId` / `@w14:textId` (Word's persistent paragraph ids) are
preserved verbatim → `Paragraph.paraId`/`.textId` and re-emitted (export de-dups so a
copy/paste clone never repeats an id; new paragraphs emit none). `@w:rsidR`/`@w:rsidP`/…
(revision-save ids) are dropped 🚫.

---

## Tables — `w:tbl`

Decoded in `documentParser.ts` (`parseTable`/`parseCell`/`parseRowProps`).

```
w:tbl
├─ w:tblGrid/w:gridCol @w:w ............... ✅ colFractions
├─ w:tblPr
│  ├─ w:tblStyle ......................... ✅
│  ├─ w:tblLook (first/last row/col, bands) ✅
│  ├─ w:tblBorders ....................... ⚠️ @val/@sz/@color per edge; @space/@themeColor dropped
│  ├─ w:shd .............................. ✅
│  ├─ w:tblCellMar ....................... ✅ (top/right/bottom/left @w:w)
│  ├─ w:tblLayout (@w:type) .............. ✅
│  ├─ w:tblW (@w:type dxa/pct, @w:w) ..... ✅
│  ├─ w:jc (table alignment) ............. ✅
│  ├─ w:tblInd (dxa) ..................... ✅
│  ├─ w:bidiVisual ....................... ✅
│  ├─ w:tblOverlap ....................... ✅
│  ├─ w:tblCaption ....................... ✅
│  ├─ w:tblDescription ................... ✅
│  ├─ w:tblCellSpacing ................... ❌
│  ├─ w:tblStyleRow/ColBandSize .......... ❌ at instance level (✅ on table *styles*)
│  ├─ w:tblpPr (floating position) ....... ❌
│  └─ w:tblHeader (on tblPr) ............. ❌
├─ w:tr → w:trPr
│  ├─ w:trHeight (@val + @hRule) ......... ✅ (auto dropped)
│  ├─ w:cantSplit ....................... ✅
│  ├─ w:tblHeader ....................... ✅
│  ├─ w:jc (row alignment) .............. ❌
│  ├─ w:tblCellSpacing .................. ❌
│  └─ w:cnfStyle ........................ ❌
└─ w:tc → w:tcPr
   ├─ w:gridSpan ........................ ✅ colSpan
   ├─ w:vMerge .......................... ✅ rowSpan
   ├─ w:tcBorders ....................... ⚠️ (as w:tblBorders — @space/@themeColor dropped)
   ├─ w:shd ............................. ✅
   ├─ w:tcMar ........................... ✅
   ├─ w:tcW (dxa/pct) ................... ✅
   ├─ w:vAlign (center/bottom) .......... ✅ (top = default)
   ├─ w:textDirection ................... ✅
   ├─ w:noWrap .......................... ✅
   ├─ w:tcFitText ....................... ✅
   ├─ w:hideMark ........................ ✅
   └─ w:cnfStyle ........................ ❌
```

---

## Sections — `w:sectPr`

Decoded in `documentParser.ts` (`parseSection`) and `props.ts` (mid-document breaks on `w:pPr/w:sectPr`).

```
w:sectPr
├─ w:pgSz (@w:w/@w:h) .................... ✅  (@w:orient dropped)
├─ w:pgMar (t/r/b/l/header/footer) ...... ✅  (@w:gutter dropped)
├─ w:headerReference / w:footerReference . ✅ default/first/even (first via titlePg, even via settings)
├─ w:titlePg ............................ ✅
├─ w:cols (@num/@space/@sep + per-col) ... ✅
├─ w:pgBorders (offsetFrom + 4 edges) .... ⚠️ @val/@sz/@space/@color; @themeColor/@frame/@display dropped
├─ w:pgNumType .......................... ⚠️ only @w:start; @fmt/@chapStyle/@chapSep dropped
├─ w:lnNumType (countBy/start/dist/restart) ✅
├─ w:type (section break kind) .......... ✅ continuous/nextPage/evenPage/oddPage
├─ w:vAlign (vertical page alignment) .... ❌
├─ w:docGrid ............................ ❌
├─ w:paperSrc ........................... ❌
├─ w:formProt ........................... ❌
├─ w:textDirection (section) ............ ❌
├─ w:bidi (section) ..................... ❌
├─ w:rtlGutter .......................... ❌
└─ w:endnotePr / w:footnotePr ........... ❌
```

> **See also #162** (open): `w:sectPr` continuation inheritance across sections.

---

## Numbering — `numbering.xml`

Decoded in `numbering.ts`.

```
w:numbering
├─ w:abstractNum (@w:abstractNumId) ...... ✅
├─ w:num (@w:numId → abstractNumId) ...... ✅
│  └─ w:lvlOverride (+ w:startOverride) .. ✅
└─ w:lvl
   ├─ w:numFmt .......................... ✅
   ├─ w:lvlText ......................... ✅
   ├─ w:start ........................... ✅
   ├─ w:ilvl ............................ ✅
   ├─ w:pPr ............................. ⚠️ only w:ind (left/start/hanging)
   ├─ w:rPr ............................. ✅ (via decodeRunProps)
   ├─ w:lvlJc ........................... ❌
   ├─ w:suff ............................ ❌
   ├─ w:isLgl ........................... ❌
   ├─ w:lvlRestart ...................... ❌
   ├─ w:pStyle (level-linked style) ..... ❌
   └─ w:legacy .......................... ❌
w:numStyleLink / w:styleLink ............ ❌ (numStyleLink-only defs skipped)
```

---

## Styles — `styles.xml`

Decoded in `styles.ts`.

```
w:styles
├─ w:docDefaults (w:rPrDefault/w:pPrDefault) ✅
└─ w:style (@styleId, @type para/char/table/num)
   ├─ w:name .......................... ✅
   ├─ w:basedOn ....................... ✅ (chains resolved, cycle-guarded)
   ├─ w:rPr / w:pPr ................... ✅
   ├─ w:default (@w:default) .......... ✅
   ├─ table: w:tblPr/w:tblBorders ..... ✅
   ├─ table: w:tblPr/w:shd ............ ✅
   ├─ table: w:tblStyleRow/ColBandSize  ✅
   ├─ w:tblStylePr (conditional bands)  ✅ @type + band rPr/pPr/tcBorders/tcShd (band trPr/tblPr dropped)
   ├─ w:next .......................... ❌
   ├─ w:link .......................... ❌
   ├─ w:uiPriority .................... ❌
   ├─ w:semiHidden / w:unhideWhenUsed . ❌
   ├─ w:qFormat ....................... ❌
   └─ w:aliases/@autoRedefine/@hidden/@locked/@rsid ❌
```

---

## Fields

Decoded in `documentParser.ts` (`handleFldChar`, `fldSimple`) + `@cw/shared` field parser.

```
Complex field (w:fldChar begin/separate/end + w:instrText) ✅
w:fldSimple (@w:instr) .................. ✅
Recognized as first-class field objects:
├─ PAGE / NUMPAGES ...................... ✅ (live {page}/{pages} tokens in headers/footers)
├─ DATE / TIME ......................... ✅
├─ IF .................................. ✅
├─ TOC ................................. ✅ (instruction + PAGEREF/HYPERLINK anchors → TOC entries)
├─ PAGEREF ............................. ✅ (anchor tagging)
└─ HYPERLINK \l ........................ ✅ (anchor extraction)
Other instructions (REF, AUTHOR, FILENAME, SEQ, STYLEREF, MERGEFIELD, …):
   → generic "custom field" IF the host registers it (isCustomFieldInstruction),
     ELSE the cached result is imported as static text.  ⚠️ not individually modeled
```

---

## Content controls — `w:sdt` (block + inline)

Decoded in `documentParser.ts` (`parseSdtPr`); nesting via `sdtPath`.

```
w:sdtPr
├─ w:alias ............................. ✅
├─ w:tag ............................... ✅
├─ w:text (→ plainText) ................ ✅
├─ w:dropDownList / w:comboBox (+listItem) ✅
├─ w:date (+ w:dateFormat) ............. ✅
├─ w14:checkbox (+ w14:checked) ........ ✅
│   └─ w14:checkedState / w14:uncheckedState ✅ (glyph font + code point preserved)
├─ w:showingPlcHdr ..................... ✅
├─ w:lock (@w:val) ..................... ✅
├─ w:id ................................ ⚠️ fresh internal id minted (original not kept)
├─ w:placeholder (docPart ref) ......... ❌ (only the showingPlcHdr flag kept)
├─ w:dataBinding ....................... ❌
├─ w:picture ........................... ❌ (treated as richText)
├─ w:group ............................. ❌ (unwrapped)
└─ w:richText/citation/bibliography/docPartObj/equation ❌ (default richText / unwrap)
```

---

## Drawing / images

Decoded in `documentParser.ts` (`parseDrawing`, `parseVmlPict`).

```
w:drawing → wp:inline / wp:anchor → a:blip @r:embed ✅
├─ a:blip @r:link (linked / "Link to File") ✅ (external URL → ImageBlock.externalSrc; export re-emits r:link + TargetMode="External"; local file: paths skipped)
├─ wp:extent (@cx/@cy) ................. ✅
├─ a:srcRect (crop l/t/r/b) ............ ✅
├─ wrap: square/tight/through → square . ✅
├─ wrap: none (floating) + positionH/V . ✅ (@relativeFrom, align/posOffset, behindDoc, decorative)
│   └─ wp14 pct position (pctPosH/VOffset) ⚠️ percentage dropped; mc:Fallback wp:posOffset recovered (#188)
├─ wp14:anchorId / wp14:editId (drawing id) ✅ ImageBlock.drawingId (preserved verbatim; export de-dups) (#188)
├─ wp14 relative size (sizeRelH/V, pctWidth/Height) ❌ (absolute wp:extent kept)
├─ wrap: topAndBottom / overlapping ..... ⚠️ demoted to inline flow (warning)
├─ mc:AlternateContent (Choice/Fallback) ✅ (run-level: Choice preferred; inside positionH/V: Fallback recovered)
└─ VML w:pict → v:shape/v:imagedata @r:id ✅ (image only)
w:object (OLE) .......................... 🚫 skipped (warning)
Charts (c:chart) ........................ ❌
VML shapes / textboxes / canvas (non-image) 🚫
```

### Drawing shapes (DrawingML wps:wsp) — issue #206

Preset drawing shapes round-trip as bare DrawingML (no VML / `mc:AlternateContent`).
Decoded in `documentParser.ts` (`parseShapeWsp`), emitted by `shapeParagraphXml`.

```
w:drawing → wp:inline → a:graphicData @uri=…/wordprocessingShape → wps:wsp ✅ ShapeBlock
├─ wps:spPr / a:prstGeom @prst .......... ✅ rect/roundRect/ellipse/triangle/diamond/right|leftArrow/line; other presets → rect + warning
├─ wp:extent (@cx/@cy) → widthPx/heightPx ✅
├─ a:solidFill / a:noFill (shape fill) .. ✅ hex color / explicit none (absent = default)
├─ a:ln (@w + a:solidFill / a:noFill) ... ✅ solid outline (point width) / explicit none
│   └─ a:prstDash (@val) ................ ✅ solid/dash/dot/dashDot/lgDash
├─ wp14:anchorId / wp14:editId .......... ✅ ShapeBlock.drawingId (preserved verbatim; export de-dups)
├─ a:avLst / a:gd (adjust handles) ...... ✅ raw guide values (name → `val N`) round-trip via geometry.adjust
├─ a:xfrm @rot (rotation) ............... ✅ ShapeBlock.rotation (degrees ↔ 60000ths)
├─ wps:txbx / w:txbxContent (text body) . ✅ ShapeBlock.text — paragraph flow, read-only render + round-trip (edit ❌ Part 6)
└─ wp:anchor (float / wrap / z-order) ... ❌ in-flow only (Part 4)
```

> Parts 2–9 (geometry & style breadth, text boxes, positioning, VML import,
> freeform/grouped shapes) are tracked in [SHAPES_PLAN.md](./SHAPES_PLAN.md).

---

> SVG images can't be decoded by pdfkit for PDF/DOCX export yet — see **#116** (open).

---

## Math (OMML — `m:` namespace)

```
m:oMath (inline) ........................ ✅ CharStyle.equation (MathML AST)
m:oMathPara (display) ................... ✅ EquationBlock, wrapped in a w:p (#193)
m:oMathParaPr/m:jc (alignment) .......... ✅ EquationBlock.align
display equation size (drag-to-resize) .. ✅ EquationBlock.scale ↔ w:p/w:pPr/w:rPr/w:sz
```

---

## Notes — `footnotes.xml` / `endnotes.xml`

```
w:footnote / w:endnote bodies ........... ✅ (separator/continuation pseudo-notes skipped)
w:footnoteReference / w:endnoteReference . ✅ (id markers)
w:footnoteRef / w:endnoteRef ............ ✅ number painted by engine (marker skipped by design)
Tables inside notes ..................... ❌ (paragraphs only, warning)
w:footnotePr / w:endnotePr (numFmt/start/restart/pos) ❌
```

---

## Bookmarks / comments / revisions

```
w:bookmarkStart / w:bookmarkEnd ......... ✅ block + inline (offsets); _GoBack/_Toc filtered
w:ins (tracked insertion) ............... ⚠️ content kept; @author/@date metadata dropped
w:del (tracked deletion) ................ ⚠️ deleted text removed (content dropped)
w:moveFrom / w:moveTo ................... 🚫 tracked moves not modeled
w:comment* / w:commentRange* ............ 🚫 comments not imported
w:smartTag .............................. ⚠️ unwrapped
w:proofErr .............................. ⚠️ ignored
w:rPrChange/w:pPrChange/w:tblPrChange … . 🚫 revision marks not modeled
```

---

## Settings & other parts

```
settings.xml
├─ w:evenAndOddHeadersAndFooters ........ ✅
├─ w:defaultTabStop .................... ✅ (non-default only)
├─ w:compat/w:compatSetting ............ ✅ (round-tripped verbatim)
├─ w15:docId (accepts w14:docId too) ... ✅ (document identity GUID, round-tripped verbatim)
└─ w:proofState/w:zoom/w:trackChanges/w:documentProtection/w:mailMerge … ❌
w:document/w:background @w:color ........ ✅ page fill (VML/gradient/image fill dropped)
theme1.xml (font/color scheme) .......... ✅ (resolution only)
Glossary (word/glossary/*) .............. ❌
Custom XML (customXml/*, w:dataBinding) . ❌
docProps/core|app|custom.xml (metadata) . ❌
webSettings.xml / fontTable.xml ......... ❌
```

---

## Cross-cutting notes

- **Borders** (`borders.ts`) decode only `@w:val`, `@w:sz`, `@w:color` per edge — `@w:space` and
  `@w:themeColor` are dropped for table/cell/paragraph/run borders (page borders additionally read
  `@w:space`). Line styles collapse to single/double/dashed/dotted (thick/wave/3-D → plain single).
- **Shading** (`decodeShdFill`) reads `@w:fill` (+ `@w:val`/`@w:color` for pattern approximation);
  other pattern types are approximated by color or cleared.
- **Intentionally out of scope** (never imported): comments, tracked-change metadata & moves, OLE
  objects, non-image VML shapes / textboxes / drawing canvas, charts, glossary document, custom XML,
  `docProps` metadata, footnote/endnote numbering properties.
- **Custom blocks** (`registerBlockType`, `CustomBlock`) have no native OOXML: `.docx` export is
  lossy by default — an empty placeholder `w:p` plus a `custom-block-dropped` warning; a registered
  type may supply `toOOXML(data)` to control its DOCX output. There is no import round-trip (a custom
  block never originates from a `.docx`). **PDF export DOES render them** — the block's `paint` is
  replayed through a Canvas2D→pdfkit vector shim, headlessly (only an unregistered type, or a
  `pdf: "raster"` block, still reserves space + warns `custom-block-not-rendered`).

---

## Actionable gaps (❌ that are modelable round-trip fidelity)

Same class as #161/#167 — candidates for future small PRs:

- Run complex-script pairs: `w:bCs`, `w:iCs`, divergent `w:szCs`; run `w:shd` (character shading); `w:noProof`.
- Border `@w:space` / `@w:themeColor`; `w:pgNumType/@w:fmt`; `w:trPr/w:jc`; `w:tblCellSpacing`.
- Numbering `w:lvlJc` / `w:suff` / `w:pStyle`; style `w:link` / `w:qFormat` / `w:uiPriority`.
- Section `w:docGrid` / `w:vAlign`; footnote/endnote numbering props.

> This list is the working backlog; promote an item to a GitHub issue before taking it on.

