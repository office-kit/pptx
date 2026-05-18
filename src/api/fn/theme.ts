// Theme + fonts.

import {
  NS,
  type XmlElement,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
} from '../../internal/xml/index.ts';
import { INTERNAL_PACKAGE, type PresentationData } from '../_internal-symbols.ts';
import { decode } from './_helpers.ts';

// ---------------------------------------------------------------------------
// Theme.

/**
 * The named color scheme on a presentation's theme. Each slot is a
 * `#RRGGBB` string — `sysClr` slots are flattened to their cached
 * `lastClr` value.
 */
export interface PresentationTheme {
  readonly name: string;
  readonly dark1: string;
  readonly light1: string;
  readonly dark2: string;
  readonly light2: string;
  readonly accent1: string;
  readonly accent2: string;
  readonly accent3: string;
  readonly accent4: string;
  readonly accent5: string;
  readonly accent6: string;
  readonly hyperlink: string;
  readonly followedHyperlink: string;
}

const THEME_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml';

const NAME_THEME_ELEMENTS = qname('a', 'themeElements', NS.dml);
const NAME_CLR_SCHEME = qname('a', 'clrScheme', NS.dml);
const NAME_SRGB_CLR = qname('a', 'srgbClr', NS.dml);
const NAME_SYS_CLR = qname('a', 'sysClr', NS.dml);

const readSchemeSlot = (parent: XmlElement, local: string): string => {
  const slot = firstChildElement(parent, qname('a', local, NS.dml));
  if (!slot) return '';
  const srgb = firstChildElement(slot, NAME_SRGB_CLR);
  if (srgb) {
    const v = getAttrValue(srgb, qname('', 'val', ''));
    if (v) return `#${v.toUpperCase()}`;
  }
  const sys = firstChildElement(slot, NAME_SYS_CLR);
  if (sys) {
    const last = getAttrValue(sys, qname('', 'lastClr', ''));
    if (last) return `#${last.toUpperCase()}`;
  }
  return '';
};

/**
 * Returns the first theme's color scheme as `#RRGGBB` strings, or
 * `null` if the package carries no theme. Each accent slot maps
 * directly to the `accent1`–`accent6` chart palette defaults.
 *
 * Multi-master decks may carry several themes — v1 surfaces only the
 * first one found (alphabetical by part name). Per-master theme
 * lookup will land if a concrete user need shows up.
 */
export const getPresentationTheme = (pres: PresentationData): PresentationTheme | null => {
  const pkg = pres[INTERNAL_PACKAGE];
  const themePart = pkg.parts
    .filter((p) => p.contentType === THEME_CONTENT_TYPE)
    .sort((a, b) => a.name.localeCompare(b.name))[0];
  if (!themePart) return null;
  const root = parseXml(decode(themePart.data)).root;
  const themeElements = firstChildElement(root, NAME_THEME_ELEMENTS);
  if (!themeElements) return null;
  const clrScheme = firstChildElement(themeElements, NAME_CLR_SCHEME);
  if (!clrScheme) return null;
  return {
    name: getAttrValue(clrScheme, qname('', 'name', '')) ?? '',
    dark1: readSchemeSlot(clrScheme, 'dk1'),
    light1: readSchemeSlot(clrScheme, 'lt1'),
    dark2: readSchemeSlot(clrScheme, 'dk2'),
    light2: readSchemeSlot(clrScheme, 'lt2'),
    accent1: readSchemeSlot(clrScheme, 'accent1'),
    accent2: readSchemeSlot(clrScheme, 'accent2'),
    accent3: readSchemeSlot(clrScheme, 'accent3'),
    accent4: readSchemeSlot(clrScheme, 'accent4'),
    accent5: readSchemeSlot(clrScheme, 'accent5'),
    accent6: readSchemeSlot(clrScheme, 'accent6'),
    hyperlink: readSchemeSlot(clrScheme, 'hlink'),
    followedHyperlink: readSchemeSlot(clrScheme, 'folHlink'),
  };
};

/**
 * The theme's font scheme, flattened to the typefaces the runs in a
 * deck inherit by default. `major*` is the heading font (slide
 * titles, chart titles); `minor*` is the body font.
 *
 * Each field carries the Latin / East-Asian / Complex-Script
 * typeface name as written on the theme. Empty values are
 * normalized to `null`.
 */
export interface PresentationFonts {
  readonly majorLatin: string | null;
  readonly majorEastAsian: string | null;
  readonly majorComplexScript: string | null;
  readonly minorLatin: string | null;
  readonly minorEastAsian: string | null;
  readonly minorComplexScript: string | null;
}

const readTypeface = (parent: XmlElement | null, local: string): string | null => {
  if (!parent) return null;
  const el = firstChildElement(parent, qname('a', local, NS.dml));
  if (!el) return null;
  const v = getAttrValue(el, qname('', 'typeface', ''));
  if (!v) return null;
  return v;
};

/**
 * Returns the first theme's font scheme, or `null` when the package
 * carries no theme. As with `getPresentationTheme`, multi-master
 * decks surface only the first theme found (alphabetical by part
 * name); per-master font lookup will land if needed.
 */
export const getPresentationFonts = (pres: PresentationData): PresentationFonts | null => {
  const pkg = pres[INTERNAL_PACKAGE];
  const themePart = pkg.parts
    .filter((p) => p.contentType === THEME_CONTENT_TYPE)
    .sort((a, b) => a.name.localeCompare(b.name))[0];
  if (!themePart) return null;
  const root = parseXml(decode(themePart.data)).root;
  const themeElements = firstChildElement(root, NAME_THEME_ELEMENTS);
  if (!themeElements) return null;
  const fontScheme = firstChildElement(themeElements, qname('a', 'fontScheme', NS.dml));
  if (!fontScheme) return null;
  const majorFont = firstChildElement(fontScheme, qname('a', 'majorFont', NS.dml));
  const minorFont = firstChildElement(fontScheme, qname('a', 'minorFont', NS.dml));
  return {
    majorLatin: readTypeface(majorFont, 'latin'),
    majorEastAsian: readTypeface(majorFont, 'ea'),
    majorComplexScript: readTypeface(majorFont, 'cs'),
    minorLatin: readTypeface(minorFont, 'latin'),
    minorEastAsian: readTypeface(minorFont, 'ea'),
    minorComplexScript: readTypeface(minorFont, 'cs'),
  };
};
