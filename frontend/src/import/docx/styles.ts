// styles.xml → StyleResolver. This is where docx fidelity lives or dies:
// most real-world formatting (fonts, sizes, heading looks) arrives through the
// style cascade, not direct formatting.
//
// Resolution order (ECMA-376 §17.7.2):
//   docDefaults → paragraph-style basedOn chain → character-style chain → direct
//
// Within a basedOn chain, properties inherit child-over-parent (plain override).
// ACROSS the hierarchy levels, *toggle* properties (bold/italic/strike/vanish)
// combine by XOR (§17.7.3): paragraph style bold + character style bold = NOT
// bold. Direct formatting is absolute and exempt from XOR. Getting this wrong
// is the classic "everything is bold" import bug.
//
// Theme indirections (w:asciiTheme, w:themeColor) are translated to concrete
// values here, after merging — any layer may carry the reference.

import { decodeBorders, decodeShdFill, mergeBorders } from "./borders";
import { decodeParaProps, decodeRunProps } from "./props";
import { applyTintShade, themeColor, themeFont, type Theme } from "./theme";
import type { IRBorders, IRParaProps, IRRunProps } from "./types";
import { WarningSink } from "./types";
import { attr, el, els, parseXml, rootEl, val } from "./xml";

export interface StyleDef {
  id: string;
  type: string; // "paragraph" | "character" | "table" | "numbering"
  /** w:name — the human-readable name ("Heading 2"); generated documents use
   *  opaque numeric styleIds, so the NAME is what the style gallery shows. */
  name?: string;
  basedOnId?: string;
  rPr: IRRunProps;
  pPr: IRParaProps;
  /** Table styles only: w:tblPr/w:tblBorders. */
  tblBorders?: IRBorders;
  /** Table styles only: w:tblPr/w:shd → CSS fill. */
  tblShd?: string;
  /** Table styles only: each w:tblStylePr conditional band, keyed by its w:type
   *  (firstRow, band1Horz, …) — these map 1:1 to the model's TableCond. */
  tblStylePr?: Record<string, TblCondDef>;
  /** Table styles only: w:tblStyleRowBandSize / w:tblStyleColBandSize. */
  rowBandSize?: number;
  colBandSize?: number;
}

/** One w:tblStylePr conditional band's captured formatting. */
export interface TblCondDef {
  rPr: IRRunProps;
  pPr: IRParaProps;
  borders?: IRBorders;
  shd?: string;
}

/** A table style's borders/shd, with its basedOn chain already collapsed. */
export interface ResolvedTableStyle {
  borders?: IRBorders;
  shd?: string;
}

export interface StylesData {
  docDefaultsRun: IRRunProps;
  docDefaultsPara: IRParaProps;
  styles: Map<string, StyleDef>;
  /** The w:default="1" paragraph style (usually "Normal") — applies to
   *  paragraphs that carry no w:pStyle at all. */
  defaultParaStyleId?: string;
}

export const EMPTY_STYLES: StylesData = { docDefaultsRun: {}, docDefaultsPara: {}, styles: new Map() };

export function parseStylesXml(xmlText: string, warnings: WarningSink): StylesData {
  const data: StylesData = { docDefaultsRun: {}, docDefaultsPara: {}, styles: new Map() };
  const root = rootEl(parseXml(xmlText, "styles.xml"), "w:styles");
  if (!root) return data;

  const docDefaults = el(root, "w:docDefaults");
  if (docDefaults) {
    const rPrDefault = el(docDefaults, "w:rPrDefault");
    const rPr = rPrDefault && el(rPrDefault, "w:rPr");
    if (rPr) data.docDefaultsRun = decodeRunProps(rPr);
    const pPrDefault = el(docDefaults, "w:pPrDefault");
    const pPr = pPrDefault && el(pPrDefault, "w:pPr");
    if (pPr) data.docDefaultsPara = decodeParaProps(pPr, warnings);
  }

  for (const style of els(root, "w:style")) {
    const id = attr(style, "w:styleId");
    const type = attr(style, "w:type");
    if (!id || !type) continue;
    const rPrEl = el(style, "w:rPr");
    const pPrEl = el(style, "w:pPr");
    const def: StyleDef = {
      id,
      type,
      rPr: rPrEl ? decodeRunProps(rPrEl) : {},
      pPr: pPrEl ? decodeParaProps(pPrEl, warnings) : {},
    };
    const name = val(style, "w:name");
    if (name) def.name = name;
    const basedOnId = val(style, "w:basedOn");
    if (basedOnId) def.basedOnId = basedOnId;
    if (type === "table") {
      const tblPr = el(style, "w:tblPr");
      if (tblPr) {
        const borders = decodeBorders(el(tblPr, "w:tblBorders"));
        if (borders) def.tblBorders = borders;
        const shd = decodeShdFill(el(tblPr, "w:shd"));
        if (shd) def.tblShd = shd;
        const rowBand = val(tblPr, "w:tblStyleRowBandSize");
        if (rowBand && Number.isFinite(Number(rowBand))) def.rowBandSize = Number(rowBand);
        const colBand = val(tblPr, "w:tblStyleColBandSize");
        if (colBand && Number.isFinite(Number(colBand))) def.colBandSize = Number(colBand);
      }
      const conds: Record<string, TblCondDef> = {};
      for (const sp of els(style, "w:tblStylePr")) {
        const condType = attr(sp, "w:type");
        if (!condType) continue;
        const rPrEl = el(sp, "w:rPr");
        const pPrEl = el(sp, "w:pPr");
        const tcPr = el(sp, "w:tcPr");
        const cond: TblCondDef = {
          rPr: rPrEl ? decodeRunProps(rPrEl) : {},
          pPr: pPrEl ? decodeParaProps(pPrEl, warnings) : {},
        };
        const cb = tcPr && decodeBorders(el(tcPr, "w:tcBorders"));
        if (cb) cond.borders = cb;
        const cs = tcPr && decodeShdFill(el(tcPr, "w:shd"));
        if (cs) cond.shd = cs;
        conds[condType] = cond;
      }
      if (Object.keys(conds).length > 0) def.tblStylePr = conds;
    }
    data.styles.set(id, def);
    const isDefault = attr(style, "w:default");
    if (type === "paragraph" && (isDefault === "1" || isDefault === "true")) {
      data.defaultParaStyleId = id;
    }
  }
  return data;
}

/** Collapse a table style's basedOn chain into one borders/shd spec
 *  (root → leaf, child overriding per edge). Cycle-guarded. */
export function resolveTableStyle(data: StylesData, styleId: string): ResolvedTableStyle {
  let borders: IRBorders | undefined;
  let shd: string | undefined;
  const seen = new Set<string>();
  const chain: StyleDef[] = [];
  for (
    let cur = data.styles.get(styleId);
    cur && !seen.has(cur.id);
    cur = cur.basedOnId ? data.styles.get(cur.basedOnId) : undefined
  ) {
    seen.add(cur.id);
    chain.unshift(cur); // root first
  }
  for (const def of chain) {
    borders = mergeBorders(borders, def.tblBorders);
    if (def.tblShd !== undefined) shd = def.tblShd;
  }
  const out: ResolvedTableStyle = {};
  if (borders) out.borders = borders;
  if (shd !== undefined) out.shd = shd;
  return out;
}

// ---------------------------------------------------------------------------
// Resolver

export interface StyleResolver {
  /** Effective run props: cascade + direct formatting, theme refs translated. */
  run(pStyleId: string | undefined, direct: IRRunProps): IRRunProps;
  /** Effective paragraph props: cascade + direct formatting. */
  para(direct: IRParaProps): IRParaProps;
  /** True when this style (or a basedOn ancestor) is a built-in heading — its
   *  w:name is "Heading 1".."Heading 9". Used to decide whether a section break
   *  starts a real new (heading-led) section, which page-breaks, vs flows. */
  isHeading(pStyleId: string | undefined): boolean;
}

/** Toggle properties per §17.7.3 — XOR across hierarchy levels.
 *  (underline is NOT a toggle: it's an enumerated value, plain override.) */
const RUN_TOGGLES = ["bold", "italic", "strikethrough", "vanish", "caps", "smallCaps"] as const;

export function createStyleResolver(data: StylesData, theme: Theme): StyleResolver {
  // basedOn chains resolved root→leaf with override semantics, memoized per style.
  const chainRunCache = new Map<string, IRRunProps>();
  const chainParaCache = new Map<string, IRParaProps>();

  const chainOf = (id: string): StyleDef[] => {
    const chain: StyleDef[] = [];
    const seen = new Set<string>(); // cycle guard (corrupt files exist)
    for (let cur = data.styles.get(id); cur && !seen.has(cur.id); cur = cur.basedOnId ? data.styles.get(cur.basedOnId) : undefined) {
      seen.add(cur.id);
      chain.unshift(cur); // root first
    }
    return chain;
  };

  const chainRun = (id: string): IRRunProps => {
    let out = chainRunCache.get(id);
    if (!out) {
      out = chainOf(id).reduce<IRRunProps>((acc, def) => mergeRun(acc, def.rPr), {});
      chainRunCache.set(id, out);
    }
    return out;
  };

  const chainPara = (id: string): IRParaProps => {
    let out = chainParaCache.get(id);
    if (!out) {
      out = chainOf(id).reduce<IRParaProps>((acc, def) => mergeProps(acc, def.pPr), {});
      chainParaCache.set(id, out);
    }
    return out;
  };

  // Style-layer run props (no direct formatting) memoized by (pStyle, rStyle) —
  // a 70-page document reuses a handful of combinations thousands of times.
  const layerCache = new Map<string, IRRunProps>();
  const styleLayerRun = (pStyleId: string | undefined, rStyleId: string | undefined): IRRunProps => {
    const key = `${pStyleId ?? ""}|${rStyleId ?? ""}`;
    let out = layerCache.get(key);
    if (out) return out;

    const p = pStyleId ? chainRun(pStyleId) : {};
    const r = rStyleId ? chainRun(rStyleId) : {};
    out = mergeRun(mergeRun(data.docDefaultsRun, p), r);

    // Toggles: XOR across levels, overriding what the plain merge produced.
    for (const t of RUN_TOGGLES) {
      if (data.docDefaultsRun[t] === undefined && p[t] === undefined && r[t] === undefined) continue;
      const base = data.docDefaultsRun[t] ?? false;
      out[t] = (base !== (p[t] === true)) !== (r[t] === true);
    }
    layerCache.set(key, out);
    return out;
  };

  return {
    run(pStyleId, direct) {
      const effectivePid = pStyleId ?? data.defaultParaStyleId;
      const layer = styleLayerRun(effectivePid, direct.styleId);
      // Direct formatting is absolute — toggles included (§17.7.3 exempts it from XOR).
      return translateTheme(mergeRun(layer, direct));
    },
    para(direct) {
      const pid = direct.styleId ?? data.defaultParaStyleId;
      const chain = pid ? chainPara(pid) : {};
      const eff = mergeProps(mergeProps(data.docDefaultsPara, chain), direct);
      // namedStyle in the model should reflect the document's own reference.
      if (direct.styleId) eff.styleId = direct.styleId;
      else delete eff.styleId;
      return eff;
    },
    isHeading(pStyleId) {
      if (!pStyleId) return false;
      // "heading" as a whole word, or immediately before a level digit — matches
      // "Heading 1" / "Heading1" (name or styleId) and named variants like "TOC
      // Heading Custom" (the TOC's own title style), but NOT "toc 1" (the entry
      // style) or "Subheading". Word's built-in headings use either name or id;
      // generated docs keep the human name on an opaque numeric id.
      const isHeadingLabel = (s: string | undefined): boolean => s !== undefined && /(^|\s)heading(\s|\d|$)/i.test(s);
      return chainOf(pStyleId).some((def) => isHeadingLabel(def.name) || isHeadingLabel(def.id));
    },
  };

  function translateTheme(props: IRRunProps): IRRunProps {
    const out = { ...props };
    if (!out.fontAscii && out.fontThemeAscii) {
      const font = themeFont(theme, out.fontThemeAscii);
      if (font) out.fontAscii = font;
    }
    if (!out.fontHAnsi && out.fontThemeHAnsi) {
      const font = themeFont(theme, out.fontThemeHAnsi);
      if (font) out.fontHAnsi = font;
    }
    if (!out.fontCs && out.fontThemeCs) {
      const font = themeFont(theme, out.fontThemeCs);
      if (font) out.fontCs = font;
    }
    if (!out.fontEastAsia && out.fontThemeEastAsia) {
      const font = themeFont(theme, out.fontThemeEastAsia);
      if (font) out.fontEastAsia = font;
    }
    if ((out.color === undefined || out.color === "auto") && out.colorTheme) {
      const base = themeColor(theme, out.colorTheme);
      // Apply w:themeTint / w:themeShade so a tinted/shaded theme color resolves to
      // its actual shade, not the flat base hex.
      if (base) out.color = applyTintShade(base, out.colorThemeTint, out.colorThemeShade);
    }
    if ((out.underlineColor === undefined || out.underlineColor === "auto") && out.underlineColorTheme) {
      const color = themeColor(theme, out.underlineColorTheme);
      if (color) out.underlineColor = color;
    }
    return out;
  }
}

// ---------------------------------------------------------------------------
// Property merging (defined fields of `add` override `base`)

/** Font and color each span two linked fields (concrete + theme ref). A layer
 *  that specifies either form of the slot replaces the whole slot — otherwise a
 *  child's concrete font would lose to a grandparent's lingering theme ref. */
function mergeRun(base: IRRunProps, add: IRRunProps): IRRunProps {
  const out: IRRunProps = { ...base };
  // Each w:rFonts slot (ascii / hAnsi / cs / eastAsia) cascades independently; a
  // layer that respecifies a slot (concrete or theme ref) replaces just that slot.
  if (add.fontAscii !== undefined || add.fontThemeAscii !== undefined) {
    delete out.fontAscii;
    delete out.fontThemeAscii;
  }
  if (add.fontHAnsi !== undefined || add.fontThemeHAnsi !== undefined) {
    delete out.fontHAnsi;
    delete out.fontThemeHAnsi;
  }
  if (add.fontCs !== undefined || add.fontThemeCs !== undefined) {
    delete out.fontCs;
    delete out.fontThemeCs;
  }
  if (add.fontEastAsia !== undefined || add.fontThemeEastAsia !== undefined) {
    delete out.fontEastAsia;
    delete out.fontThemeEastAsia;
  }
  if (add.color !== undefined || add.colorTheme !== undefined) {
    delete out.color;
    delete out.colorTheme;
    delete out.colorThemeTint; // tint/shade belong to the color slot — drop with it
    delete out.colorThemeShade;
  }
  // A w:u in `add` re-declares the WHOLE underline (style + color), so a child's
  // bare underline mustn't inherit a grandparent's lingering underline color.
  if (add.underline !== undefined || add.underlineStyle !== undefined || add.underlineColor !== undefined || add.underlineColorTheme !== undefined) {
    delete out.underlineStyle;
    delete out.underlineColor;
    delete out.underlineColorTheme;
  }
  return mergeProps(out, add);
}

function mergeProps<T extends object>(base: T, add: T): T {
  const out = { ...base };
  for (const [k, v] of Object.entries(add)) {
    if (v !== undefined) (out as Record<string, unknown>)[k] = v;
  }
  return out;
}
