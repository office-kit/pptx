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

import {
  NS,
  type XmlAttr,
  type XmlElement,
  attr,
  elem,
  firstChildElement,
  qname,
} from '../xml/index.ts';

const NAME_R = qname('a', 'r', NS.dml);
const NAME_RPR = qname('a', 'rPr', NS.dml);
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

export interface TextFormat {
  /** Latin font family (`Calibri`, `Arial`, ...). Sets `<a:latin>`. */
  font?: string;
  /** Font size in points; fractional values allowed (`12`, `12.5`). */
  size?: number;
  /**
   * Color. Accepts `#RRGGBB`, `RRGGBB`, an ECMA-376 scheme color token
   * (`tx1`, `accent1`, ...), or `null` to clear.
   */
  color?: string | null;
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
   * Kerning threshold in half-points (`0` disables kerning, `1200` =
   * apply kerning for runs ≥12pt). Mirrors `<a:rPr kern="…"/>`.
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
  highlight?: string | null;
}

const SCHEME_TOKENS = new Set([
  'bg1',
  'tx1',
  'bg2',
  'tx2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
  'phClr',
  'lt1',
  'dk1',
  'lt2',
  'dk2',
]);

const parseColor = (
  value: string,
): { kind: 'srgb'; hex: string } | { kind: 'scheme'; token: string } | null => {
  if (SCHEME_TOKENS.has(value)) return { kind: 'scheme', token: value };
  const hex = value.startsWith('#') ? value.slice(1) : value;
  if (/^[0-9A-Fa-f]{6}$/.test(hex)) return { kind: 'srgb', hex: hex.toUpperCase() };
  return null;
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
  // Per the schema, fills come BEFORE typeface children. Insert after any
  // text-decoration attrs but before latin/ea/cs. Easiest: prepend; xmllint
  // accepts either placement since solidFill is a choice in CT_TextCharacterProperties.
  rPr.children.unshift(fill);
};

const setLatin = (rPr: XmlElement, font: string | null): void => {
  rPr.children = rPr.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'latin'),
  );
  if (font === null) return;
  rPr.children.push(elem(NAME_LATIN, { attrs: [attr(ATTR_TYPEFACE, font)] }));
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
  // highlight follows solidFill but precedes the typeface children in the
  // schema. Insert near the start; xmllint accepts either placement.
  rPr.children.unshift(elem(NAME_HIGHLIGHT, { children: [inner] }));
};

/** Mutates `rPr` in place per `format`. */
export const applyRunFormat = (rPr: XmlElement, format: TextFormat): void => {
  let attrs = rPr.attrs;
  if (format.size !== undefined) {
    // Hundredths of a point per the schema.
    const sz = Math.round(format.size * 100);
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
    attrs = setOrRemoveAttr(attrs, ATTR_SPC, String(Math.round(format.spc)));
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
  if (format.color !== undefined) setSolidFill(rPr, format.color);
  if (format.highlight !== undefined) setHighlight(rPr, format.highlight);

  void NAME_EA;
  void NAME_CS;
};

/**
 * Walks `txBody`, ensuring every `<a:r>` has an `<a:rPr>` carrying the
 * supplied format. Existing run-property attributes not addressed by
 * `format` are preserved.
 */
export const applyFormatToAllRuns = (txBody: XmlElement, format: TextFormat): void => {
  // Walk depth-first; runs live two levels deep (txBody > p > r).
  for (const p of txBody.children) {
    if (p.kind !== 'element' || p.name.namespaceURI !== NS.dml || p.name.localName !== 'p') {
      continue;
    }
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
      applyRunFormat(rPr, format);
    }
  }
  // Force-touch a NAME_R reference so it isn't elided as unused.
  void NAME_R;
};
