// Run-property authoring for `<a:r><a:rPr>...</a:rPr></a:r>`.
//
// The ECMA-376 surface for `CT_TextCharacterProperties` (§17.18.83) is
// huge. We expose the subset that 90% of authoring scripts reach for:
// font face, size, color, bold, italic, underline.
//
// Sizes are pt and accept fractions; serialized as hundredths of a point
// per the schema (`sz="2400"` = 24pt). Colors accept any of:
//
//   - `#RRGGBB` / `RRGGBB` for srgb
//   - `'tx1' | 'tx2' | 'bg1' | 'bg2' | 'accent1'…'accent6'` for theme
//   - `'transparent' | null` to clear
//
// Mutation strategy: walk the `txBody`, ensure each `a:r` has an `a:rPr`,
// then set the relevant attributes / child elements. Existing properties
// not addressed by the format object are preserved.

import type { Color } from './color.ts';
import {
  NS,
  type XmlAttr,
  type XmlElement,
  attr,
  elem,
  firstChildElement,
  insertChildByRank,
  qname,
} from '../xml/index.ts';
import { UNDERLINES, STRIKES } from '../enum-values.ts';
import { oneOf, fontSizeHundredthPt, textPointSpacing } from '../bounds.ts';
import { asColor, parseColor } from './color.ts';
import {
  type EffectPlacement,
  type GlowOptions,
  type ShadowOptions,
  removeEffect,
  setGlow,
  setShadow,
} from './effects.ts';
import { applySolidStroke } from './stroke.ts';

const NAME_R = qname('a', 'r', NS.dml);
const NAME_RPR = qname('a', 'rPr', NS.dml);
const NAME_LN = qname('a', 'ln', NS.dml);
const NAME_LATIN = qname('a', 'latin', NS.dml);
const NAME_EA = qname('a', 'ea', NS.dml);
const NAME_CS = qname('a', 'cs', NS.dml);
const NAME_SOLID_FILL = qname('a', 'solidFill', NS.dml);
const NAME_SRGB_CLR = qname('a', 'srgbClr', NS.dml);
const NAME_SCHEME_CLR = qname('a', 'schemeClr', NS.dml);
const ATTR_SZ = qname('', 'sz', '');
const ATTR_B = qname('', 'b', '');
const ATTR_I = qname('', 'i', '');
const ATTR_U = qname('', 'u', '');
const ATTR_STRIKE = qname('', 'strike', '');
const ATTR_SPC = qname('', 'spc', '');
const ATTR_KERN = qname('', 'kern', '');
const ATTR_BASELINE = qname('', 'baseline', '');
const ATTR_CAP = qname('', 'cap', '');
const ATTR_TYPEFACE = qname('', 'typeface', '');
const ATTR_VAL = qname('', 'val', '');
const NAME_HIGHLIGHT = qname('a', 'highlight', NS.dml);

// CT_TextCharacterProperties (a:rPr) is an xsd:sequence: children must appear
// in this order or the run fails dml/pml schema validation. Setters strip the
// existing element then re-insert at the mandated slot via insertChildByRank.
const RPR_CHILD_RANK: Record<string, number> = {
  ln: 0,
  noFill: 1,
  solidFill: 1,
  gradFill: 1,
  blipFill: 1,
  pattFill: 1,
  grpFill: 1,
  effectLst: 2,
  effectDag: 2,
  highlight: 3,
  uLnTx: 4,
  uLn: 4,
  uFillTx: 5,
  uFill: 5,
  latin: 6,
  ea: 7,
  cs: 8,
  sym: 9,
  hlinkClick: 10,
  hlinkMouseOver: 11,
  rtl: 12,
  extLst: 13,
};
const rprChildRank = (el: XmlElement): number =>
  el.name.namespaceURI === NS.dml ? (RPR_CHILD_RANK[el.name.localName] ?? 99) : 99;

export interface TextFormat {
  /** Latin font family (`Calibri`, `Arial`, ...). Sets `<a:latin>`. */
  font?: string;
  /**
   * East Asian font family (`游明朝`, `メイリオ`, ...). Sets `<a:ea>`.
   * Renderers pick this typeface for CJK glyphs independently of `font`
   * (which only governs Latin glyphs) — set both when a run mixes Latin
   * and CJK text and needs a consistent look across the whole run.
   */
  fontEastAsian?: string;
  /**
   * Complex-script font family. Sets `<a:cs>`, the typeface renderers pick
   * for complex scripts such as Arabic, Hebrew and Thai — independently of
   * `font` (Latin) and `fontEastAsian` (CJK).
   */
  fontComplexScript?: string;
  /** Font size in points; fractional values allowed (`12`, `12.5`). */
  size?: number;
  /**
   * Color. Accepts `#RRGGBB`, `RRGGBB`, an ECMA-376 scheme color token
   * (`tx1`, `accent1`, ...), or `null` to clear.
   */
  color?: Color | null;
  bold?: boolean;
  italic?: boolean;
  /**
   * Underline style. `true` is shorthand for `'sng'` (single). Pass the
   * exact `ST_TextUnderlineType` token for other styles (`'dbl'`, `'wavy'`,
   * `'dash'`, ...).
   */
  underline?: boolean | string;
  /**
   * Strikethrough style. `true` is shorthand for `'sngStrike'` (single
   * line). Pass the exact `ST_TextStrikeType` token (`'sngStrike'`,
   * `'dblStrike'`, `'noStrike'`) for other styles. `false` clears.
   */
  strike?: boolean | string;
  /**
   * Character spacing in 1/100 points (`0` = default). Negative values
   * tighten, positive values loosen. Mirrors `<a:rPr spc="…"/>`.
   */
  spc?: number;
  /**
   * Kerning threshold in 1/100 points (`ST_TextNonNegativePoint`, the same
   * unit as `spc`): `0` disables kerning, `1200` = apply kerning for runs
   * ≥12pt. Mirrors `<a:rPr kern="…"/>`.
   */
  kern?: number;
  /**
   * Baseline offset as a fraction of 1 (`0.3` = superscript ~30% up,
   * `-0.25` = subscript). PowerPoint emits ST_Percentage; this getter
   * returns the unit-fraction form for ergonomic comparisons.
   */
  baseline?: number;
  /**
   * Capitalization mode: `'none'`, `'small'` (smallCaps), or `'all'`
   * (allCaps). Mirrors `<a:rPr cap="…"/>`.
   */
  cap?: 'none' | 'small' | 'all';
  /**
   * Highlight color (cell-fill style background per run). Same color
   * format as `color`. Mirrors `<a:rPr><a:highlight>…</a:highlight></a:rPr>`.
   */
  highlight?: Color | null;
  /**
   * Outline drawn around the glyphs — `<a:rPr><a:ln>`, the character-level
   * twin of `setShapeStroke`. `null` removes it, which is not the same as
   * an outline of width 0: removing restores what the run inherits.
   */
  outline?: TextOutline | null;
  /**
   * Drop shadow behind the glyphs — `<a:outerShdw>` in the run's own
   * `<a:effectLst>`. `null` removes it.
   */
  shadow?: ShadowOptions | null;
  /**
   * Glow around the glyphs — `<a:glow>` in the run's own `<a:effectLst>`.
   * `null` removes it.
   */
  glow?: GlowOptions | null;
}

/** A run's outline: `CT_LineProperties` as far as text uses it. */
export interface TextOutline {
  /** Same accepted forms as `TextFormat.color`. */
  readonly color?: Color;
  /** Line width in EMU. PowerPoint's thinnest visible text outline is 9525 (0.75pt). */
  readonly widthEmu?: number;
}

/**
 * A run format read back from a deck. `color` and `highlight` widen to
 * `string`: when no theme is supplied, or a token is not in the scheme, the
 * readers surface the raw `<a:schemeClr val>` token as-is.
 */
export type ReadTextFormat = Omit<
  TextFormat,
  'color' | 'highlight' | 'outline' | 'shadow' | 'glow'
> & {
  color?: string | null;
  highlight?: string | null;
  outline?: ReadTextOutline | null;
  shadow?: (Omit<ShadowOptions, 'color'> & { readonly color?: string }) | null;
  glow?: (Omit<GlowOptions, 'color'> & { readonly color: string }) | null;
};

/** A run outline read back from a deck. `color` widens for the same reason. */
export type ReadTextOutline = Omit<TextOutline, 'color'> & { readonly color?: string };

/**
 * Converts a format read back from a deck into one a writer accepts. The
 * readers widen every color to `string`, because a deck can hold a scheme
 * token that is not in its theme; this checks each one and drops the property
 * (or, for a glow, the whole effect) when the writer would reject it, so the
 * round trip never writes a color the schema has no room for.
 */
export const toWritableTextFormat = (format: ReadTextFormat): TextFormat => {
  const { color, highlight, outline, shadow, glow, ...rest } = format;
  const outlineColor = outline?.color === undefined ? null : asColor(outline.color);
  const shadowColor = shadow?.color === undefined ? null : asColor(shadow.color);
  const glowColor = glow == null ? null : asColor(glow.color);
  return {
    ...rest,
    ...(color == null ? {} : { color: asColor(color) }),
    ...(highlight == null ? {} : { highlight: asColor(highlight) }),
    ...(outline == null
      ? {}
      : {
          outline: {
            ...(outline.widthEmu === undefined ? {} : { widthEmu: outline.widthEmu }),
            ...(outlineColor === null ? {} : { color: outlineColor }),
          },
        }),
    ...(shadow == null
      ? {}
      : {
          shadow: {
            ...(shadow.blurEmu === undefined ? {} : { blurEmu: shadow.blurEmu }),
            ...(shadow.offsetEmu === undefined ? {} : { offsetEmu: shadow.offsetEmu }),
            ...(shadow.angleDeg === undefined ? {} : { angleDeg: shadow.angleDeg }),
            ...(shadow.opacity === undefined ? {} : { opacity: shadow.opacity }),
            ...(shadowColor === null ? {} : { color: shadowColor }),
          },
        }),
    ...(glow == null || glowColor === null
      ? {}
      : {
          glow: {
            color: glowColor,
            ...(glow.radiusEmu === undefined ? {} : { radiusEmu: glow.radiusEmu }),
            ...(glow.opacity === undefined ? {} : { opacity: glow.opacity }),
          },
        }),
  };
};

const setOrRemoveAttr = (
  attrs: XmlAttr[],
  name: ReturnType<typeof qname>,
  value: string | null,
): XmlAttr[] => {
  const filtered = attrs.filter((a) => a.name.localName !== name.localName);
  if (value !== null) filtered.push(attr(name, value));
  return filtered;
};

const setSolidFill = (rPr: XmlElement, value: string | null): void => {
  // Remove any existing solidFill first.
  rPr.children = rPr.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'solidFill'),
  );
  if (value === null) return;
  const parsed = parseColor(value);
  if (parsed === null) throw new Error(`unrecognized color: ${value}`);
  const inner =
    parsed.kind === 'srgb'
      ? elem(NAME_SRGB_CLR, { attrs: [attr(ATTR_VAL, parsed.hex)] })
      : elem(NAME_SCHEME_CLR, { attrs: [attr(ATTR_VAL, parsed.token)] });
  const fill = elem(NAME_SOLID_FILL, { children: [inner] });
  insertChildByRank(rPr, fill, rprChildRank);
};

const setLatin = (rPr: XmlElement, font: string | null): void => {
  rPr.children = rPr.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'latin'),
  );
  if (font === null) return;
  insertChildByRank(rPr, elem(NAME_LATIN, { attrs: [attr(ATTR_TYPEFACE, font)] }), rprChildRank);
};

const setEastAsian = (rPr: XmlElement, font: string | null): void => {
  rPr.children = rPr.children.filter(
    (c) => !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'ea'),
  );
  if (font === null) return;
  insertChildByRank(rPr, elem(NAME_EA, { attrs: [attr(ATTR_TYPEFACE, font)] }), rprChildRank);
};

const setComplexScript = (rPr: XmlElement, font: string | null): void => {
  rPr.children = rPr.children.filter(
    (c) => !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'cs'),
  );
  if (font === null) return;
  insertChildByRank(rPr, elem(NAME_CS, { attrs: [attr(ATTR_TYPEFACE, font)] }), rprChildRank);
};

const setHighlight = (rPr: XmlElement, value: string | null): void => {
  rPr.children = rPr.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'highlight'),
  );
  if (value === null) return;
  const parsed = parseColor(value);
  if (parsed === null) throw new Error(`unrecognized highlight color: ${value}`);
  const inner =
    parsed.kind === 'srgb'
      ? elem(NAME_SRGB_CLR, { attrs: [attr(ATTR_VAL, parsed.hex)] })
      : elem(NAME_SCHEME_CLR, { attrs: [attr(ATTR_VAL, parsed.token)] });
  insertChildByRank(rPr, elem(NAME_HIGHLIGHT, { children: [inner] }), rprChildRank);
};

export const validateFormatEnums = (format: TextFormat, caller: string): void => {
  if (format.underline != null && typeof format.underline !== 'boolean')
    oneOf(format.underline, UNDERLINES, `${caller}: underline`);
  if (format.strike != null && typeof format.strike !== 'boolean')
    oneOf(format.strike, STRIKES, `${caller}: strike`);
  if (format.cap != null) oneOf(format.cap, ['none', 'small', 'all'], `${caller}: cap`);
};

/** Mutates `rPr` in place per `format`. */
export const applyRunFormat = (
  rPr: XmlElement,
  format: TextFormat,
  caller = 'setShapeRunFormat',
): void => {
  validateFormatEnums(format, caller);
  applyValidatedRunFormat(rPr, format);
};

const applyValidatedRunFormat = (rPr: XmlElement, format: TextFormat): void => {
  let attrs = rPr.attrs;
  if (format.size !== undefined) {
    // Hundredths of a point per the schema (ST_TextFontSize: 1..4000 pt).
    const sz = fontSizeHundredthPt(format.size * 100, 'setShapeRunFormat: size');
    attrs = setOrRemoveAttr(attrs, ATTR_SZ, String(sz));
  }
  if (format.bold !== undefined) {
    attrs = setOrRemoveAttr(attrs, ATTR_B, format.bold ? '1' : '0');
  }
  if (format.italic !== undefined) {
    attrs = setOrRemoveAttr(attrs, ATTR_I, format.italic ? '1' : '0');
  }
  if (format.underline !== undefined) {
    const value =
      format.underline === false ? 'none' : format.underline === true ? 'sng' : format.underline;
    attrs = setOrRemoveAttr(attrs, ATTR_U, value);
  }
  if (format.strike !== undefined) {
    const value =
      format.strike === false ? 'noStrike' : format.strike === true ? 'sngStrike' : format.strike;
    attrs = setOrRemoveAttr(attrs, ATTR_STRIKE, value);
  }
  if (format.spc !== undefined) {
    const spc = textPointSpacing(format.spc, 'setShapeRunFormat: spc');
    attrs = setOrRemoveAttr(attrs, ATTR_SPC, String(spc));
  }
  if (format.kern !== undefined) {
    attrs = setOrRemoveAttr(attrs, ATTR_KERN, String(Math.round(format.kern)));
  }
  if (format.baseline !== undefined) {
    // ST_Percentage; we accept the unit-fraction form on the public API
    // and serialize as the on-the-wire hundredths-of-percent integer.
    const pct = Math.round(format.baseline * 100000);
    attrs = setOrRemoveAttr(attrs, ATTR_BASELINE, String(pct));
  }
  if (format.cap !== undefined) {
    attrs = setOrRemoveAttr(attrs, ATTR_CAP, format.cap);
  }
  rPr.attrs = attrs;

  if (format.font !== undefined) setLatin(rPr, format.font);
  if (format.fontEastAsian !== undefined) setEastAsian(rPr, format.fontEastAsian);
  if (format.fontComplexScript !== undefined) setComplexScript(rPr, format.fontComplexScript);
  if (format.color !== undefined) setSolidFill(rPr, format.color);
  if (format.highlight !== undefined) setHighlight(rPr, format.highlight);
  if (format.outline !== undefined) setRunOutline(rPr, format.outline);
  if (format.shadow !== undefined) {
    if (format.shadow === null) removeEffect(rPr, 'outerShdw');
    else setShadow(rPr, format.shadow, rPrEffectPlacement);
  }
  if (format.glow !== undefined) {
    if (format.glow === null) removeEffect(rPr, 'glow');
    else setGlow(rPr, format.glow, rPrEffectPlacement);
  }
};

// `<a:effectLst>` is the third slot of CT_TextCharacterProperties, so it goes
// ahead of the first child that outranks it rather than at the end the way it
// does on `<p:spPr>`.
const rPrEffectPlacement: EffectPlacement = (rPr) => {
  const own = RPR_CHILD_RANK.effectLst!;
  for (let i = 0; i < rPr.children.length; i++) {
    const c = rPr.children[i];
    if (c?.kind === 'element' && rprChildRank(c) > own) return i;
  }
  return rPr.children.length;
};

const setRunOutline = (rPr: XmlElement, outline: TextOutline | null): void => {
  const existing = firstChildElement(rPr, NAME_LN);
  if (outline === null) {
    if (existing) rPr.children = rPr.children.filter((c) => c !== existing);
    return;
  }
  const ln = existing ?? elem(NAME_LN);
  applySolidStroke(ln, outline);
  if (!existing) insertChildByRank(rPr, ln, rprChildRank);
};

const VISUAL_RUN_ATTRIBUTES = new Set([
  'kumimoji',
  'sz',
  'b',
  'i',
  'u',
  'strike',
  'kern',
  'cap',
  'spc',
  'normalizeH',
  'baseline',
]);
const VISUAL_RUN_CHILDREN = new Set([
  'ln',
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
  'effectLst',
  'effectDag',
  'highlight',
  'uLnTx',
  'uLn',
  'uFillTx',
  'uFill',
  'latin',
  'ea',
  'cs',
  'sym',
]);

/** Retain links, language, proofing and unknown extensions when clearing appearance. */
export const resetRunFormat = (properties: XmlElement): void => {
  properties.attrs = properties.attrs.filter(
    (a) => a.name.namespaceURI !== '' || !VISUAL_RUN_ATTRIBUTES.has(a.name.localName),
  );
  properties.children = properties.children.filter(
    (c) =>
      c.kind !== 'element' ||
      c.name.namespaceURI !== NS.dml ||
      !VISUAL_RUN_CHILDREN.has(c.name.localName),
  );
};

const resetTextBodyRunFormats = (node: XmlElement): void => {
  if (
    node.name.namespaceURI === NS.dml &&
    ['rPr', 'defRPr', 'endParaRPr'].includes(node.name.localName)
  ) {
    resetRunFormat(node);
    return;
  }
  for (const child of node.children) {
    if (
      child.kind === 'element' &&
      child.name.namespaceURI === NS.dml &&
      child.name.localName !== 'extLst'
    )
      resetTextBodyRunFormats(child);
  }
};

// ECMA-376 CT_TextBodyProperties / CT_TextParagraphProperties. Keep paragraph
// level (the outline structure), language, hyperlinks and unknown extensions.
const BODY_FORMAT_ATTRIBUTES = new Set([
  'rot',
  'spcFirstLastPara',
  'vertOverflow',
  'horzOverflow',
  'vert',
  'wrap',
  'lIns',
  'tIns',
  'rIns',
  'bIns',
  'numCol',
  'spcCol',
  'rtlCol',
  'fromWordArt',
  'anchor',
  'anchorCtr',
  'forceAA',
  'upright',
  'compatLnSpc',
]);
const BODY_FORMAT_CHILDREN = new Set([
  'prstTxWarp',
  'noAutofit',
  'normAutofit',
  'spAutoFit',
  'scene3d',
  'sp3d',
  'flatTx',
]);
const PARAGRAPH_FORMAT_ATTRIBUTES = new Set([
  'marL',
  'marR',
  'indent',
  'algn',
  'defTabSz',
  'rtl',
  'eaLnBrk',
  'fontAlgn',
  'latinLnBrk',
  'hangingPunct',
]);
const PARAGRAPH_FORMAT_CHILDREN = new Set([
  'lnSpc',
  'spcBef',
  'spcAft',
  'buClrTx',
  'buClr',
  'buSzTx',
  'buSzPct',
  'buSzPts',
  'buFontTx',
  'buFont',
  'buNone',
  'buAutoNum',
  'buChar',
  'buBlip',
  'tabLst',
]);

/** Clear direct text appearance so a placeholder can inherit its layout again. */
export const resetTextBodyFormatting = (body: XmlElement): void => {
  resetTextBodyRunFormats(body);
  const visit = (node: XmlElement): void => {
    const local = node.name.localName;
    const isBody = local === 'bodyPr';
    if (isBody || local === 'pPr' || local === 'defPPr' || /^lvl[1-9]pPr$/.test(local)) {
      const attributes = isBody ? BODY_FORMAT_ATTRIBUTES : PARAGRAPH_FORMAT_ATTRIBUTES;
      const children = isBody ? BODY_FORMAT_CHILDREN : PARAGRAPH_FORMAT_CHILDREN;
      node.attrs = node.attrs.filter(
        (a) => a.name.namespaceURI !== '' || !attributes.has(a.name.localName),
      );
      node.children = node.children.filter(
        (c) =>
          c.kind !== 'element' || c.name.namespaceURI !== NS.dml || !children.has(c.name.localName),
      );
    }
    for (const child of node.children) {
      if (
        child.kind === 'element' &&
        child.name.namespaceURI === NS.dml &&
        child.name.localName !== 'extLst'
      )
        visit(child);
    }
  };
  visit(body);
};

/**
 * Walks `txBody`, ensuring every `<a:r>` has an `<a:rPr>` carrying the
 * supplied format. Existing run-property attributes not addressed by
 * `format` are preserved.
 */
export const applyFormatToAllRuns = (
  txBody: XmlElement,
  format: TextFormat,
  caller = 'setShapeTextFormat',
  reset = false,
): void => {
  validateFormatEnums(format, caller);
  applyValidatedFormatToAllRuns(txBody, format, reset);
};

export const applyValidatedFormatToAllRuns = (
  txBody: XmlElement,
  format: TextFormat,
  reset = false,
): void => {
  if (reset) {
    applyValidatedRunFormat(elem(NAME_RPR), format);
    resetTextBodyRunFormats(txBody);
  }
  // Walk depth-first; runs live two levels deep (txBody > p > r).
  for (const p of txBody.children) {
    if (p.kind !== 'element' || p.name.namespaceURI !== NS.dml || p.name.localName !== 'p') {
      continue;
    }
    // The paragraph end mark supplies the format for typing into an empty paragraph.
    const existingEnd = firstChildElement(p, qname('a', 'endParaRPr', NS.dml));
    const end = existingEnd ?? elem(qname('a', 'endParaRPr', NS.dml));
    applyValidatedRunFormat(end, format);
    if (!existingEnd && (end.attrs.length || end.children.length)) p.children.push(end);
    for (const r of p.children) {
      if (r.kind !== 'element' || r.name.namespaceURI !== NS.dml || r.name.localName !== 'r') {
        continue;
      }
      let rPr = firstChildElement(r, NAME_RPR);
      if (rPr === null) {
        rPr = elem(NAME_RPR);
        // rPr must be the first child of the run per the schema.
        r.children.unshift(rPr);
      }
      applyValidatedRunFormat(rPr, format);
    }
  }
  // Force-touch a NAME_R reference so it isn't elided as unused.
  void NAME_R;
};
