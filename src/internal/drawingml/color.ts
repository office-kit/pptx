// Shared color parsing for the authoring APIs.
//
// Accepts the three forms PowerPoint emits: srgb hex (`#RRGGBB`,
// `RRGGBB`), scheme tokens (`tx1`, `accent1`...), and an explicit `null`
// to indicate "clear / no fill". Anything else throws so callers don't
// silently emit `<a:srgbClr val="undefined"/>`.

import { type XmlElement, NS, attr, elem, qname } from '../xml/index.ts';

const NAME_SRGB_CLR = qname('a', 'srgbClr', NS.dml);
const NAME_SCHEME_CLR = qname('a', 'schemeClr', NS.dml);
const ATTR_VAL = qname('', 'val', '');

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

export type ParsedColor = { kind: 'srgb'; hex: string } | { kind: 'scheme'; token: string };

/**
 * Parses a user-supplied color string. Returns null on unrecognized input
 * so callers can decide whether to throw with a specific message.
 */
export const parseColor = (value: string): ParsedColor | null => {
  if (SCHEME_TOKENS.has(value)) return { kind: 'scheme', token: value };
  const hex = value.startsWith('#') ? value.slice(1) : value;
  if (/^[0-9A-Fa-f]{6}$/.test(hex)) return { kind: 'srgb', hex: hex.toUpperCase() };
  return null;
};

/**
 * Returns the `<a:srgbClr>` or `<a:schemeClr>` element for `value`.
 * Throws on unrecognized colors.
 */
export const buildColorElement = (value: string): XmlElement => {
  const parsed = parseColor(value);
  if (parsed === null) throw new Error(`unrecognized color: ${value}`);
  return parsed.kind === 'srgb'
    ? elem(NAME_SRGB_CLR, { attrs: [attr(ATTR_VAL, parsed.hex)] })
    : elem(NAME_SCHEME_CLR, { attrs: [attr(ATTR_VAL, parsed.token)] });
};
