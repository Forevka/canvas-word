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

import { decodeParaProps, decodeRunProps } from "./props";
import { themeColor, themeFont, type Theme } from "./theme";
import type { IRParaProps, IRRunProps } from "./types";
import { WarningSink } from "./types";
import { attr, el, els, parseXml, rootEl, val } from "./xml";

interface StyleDef {
  id: string;
  type: string; // "paragraph" | "character" | "table" | "numbering"
  basedOnId?: string;
  rPr: IRRunProps;
  pPr: IRParaProps;
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
    const basedOnId = val(style, "w:basedOn");
    if (basedOnId) def.basedOnId = basedOnId;
    data.styles.set(id, def);
    const isDefault = attr(style, "w:default");
    if (type === "paragraph" && (isDefault === "1" || isDefault === "true")) {
      data.defaultParaStyleId = id;
    }
  }
  return data;
}

// ---------------------------------------------------------------------------
// Resolver

export interface StyleResolver {
  /** Effective run props: cascade + direct formatting, theme refs translated. */
  run(pStyleId: string | undefined, direct: IRRunProps): IRRunProps;
  /** Effective paragraph props: cascade + direct formatting. */
  para(direct: IRParaProps): IRParaProps;
}

/** Toggle properties per §17.7.3 — XOR across hierarchy levels.
 *  (underline is NOT a toggle: it's an enumerated value, plain override.) */
const RUN_TOGGLES = ["bold", "italic", "strikethrough", "vanish"] as const;

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
  };

  function translateTheme(props: IRRunProps): IRRunProps {
    const out = { ...props };
    if (!out.fontAscii && out.fontThemeAscii) {
      const font = themeFont(theme, out.fontThemeAscii);
      if (font) out.fontAscii = font;
    }
    if ((out.color === undefined || out.color === "auto") && out.colorTheme) {
      const color = themeColor(theme, out.colorTheme);
      if (color) out.color = color;
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
  if (add.fontAscii !== undefined || add.fontThemeAscii !== undefined) {
    delete out.fontAscii;
    delete out.fontThemeAscii;
  }
  if (add.color !== undefined || add.colorTheme !== undefined) {
    delete out.color;
    delete out.colorTheme;
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
