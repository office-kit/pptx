// Shape effects — `<a:effectLst>` builders.
//
// Covers the most-used PowerPoint effects: outer shadow + glow. The
// element ordering on `<p:spPr>` is `xfrm → geometry → fill → ln →
// effectLst → scene3d → sp3d → extLst`. Callers locate the right
// insertion slot using `effectInsertionIndex`.

import { emuExtent } from '../bounds.ts';
import { NS, type XmlElement, attr, elem, qname } from '../xml/index.ts';
import { buildColorElement } from './color.ts';

const NAME_EFFECT_LST = qname('a', 'effectLst', NS.dml);
const NAME_OUTER_SHDW = qname('a', 'outerShdw', NS.dml);
const NAME_GLOW = qname('a', 'glow', NS.dml);
const NAME_ALPHA = qname('a', 'alpha', NS.dml);

const ATTR_BLUR_RAD = qname('', 'blurRad', '');
const ATTR_DIST = qname('', 'dist', '');
const ATTR_DIR = qname('', 'dir', '');
const ATTR_ALGN = qname('', 'algn', '');
const ATTR_ROT_WITH_SHAPE = qname('', 'rotWithShape', '');
const ATTR_RAD = qname('', 'rad', '');
const ATTR_VAL = qname('', 'val', '');

export interface ShadowOptions {
  /** `#RRGGBB`, bare `RRGGBB`, or scheme token. Defaults to black. */
  readonly color?: string;
  /** Edge blur in EMU. Defaults to 50800 (4pt). */
  readonly blurEmu?: number;
  /** Offset distance in EMU. Defaults to 38100 (3pt). */
  readonly offsetEmu?: number;
  /**
   * Direction in degrees, measured clockwise from the right (3 o'clock).
   * Defaults to 45° (down-right).
   */
  readonly angleDeg?: number;
  /** Opacity (0–1). Defaults to fully opaque. */
  readonly opacity?: number;
}

export interface GlowOptions {
  /** `#RRGGBB`, bare `RRGGBB`, or scheme token. */
  readonly color: string;
  /** Glow radius in EMU. Defaults to 63500 (5pt). */
  readonly radiusEmu?: number;
  /** Opacity (0–1). Defaults to fully opaque. */
  readonly opacity?: number;
}

/**
 * Where an `<a:effectLst>` goes inside its host. `<p:spPr>` is the default;
 * `<a:rPr>` has its own order, so the run-level setters pass their own.
 */
export type EffectPlacement = (host: XmlElement) => number;

/**
 * Computes the index inside `host.children` where an `<a:effectLst>`
 * should be inserted to satisfy the spec's child ordering on
 * `<p:spPr>`.
 */
const effectInsertionIndex: EffectPlacement = (host: XmlElement): number => {
  for (let i = 0; i < host.children.length; i++) {
    const c = host.children[i];
    if (c?.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (c.name.localName === 'scene3d' || c.name.localName === 'sp3d') return i;
    if (c.name.localName === 'extLst') return i;
  }
  return host.children.length;
};

const removeEffectLst = (host: XmlElement): void => {
  host.children = host.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'effectLst'),
  );
};

// CT_EffectList is a sequence, not a choice (dml-main.xsd §20.1.8.25): each
// effect appears at most once, in this order. Writing one out of order, or
// twice, is schema-invalid.
const EFFECT_ORDER = [
  'blur',
  'fillOverlay',
  'glow',
  'innerShdw',
  'outerShdw',
  'prstShdw',
  'reflection',
  'softEdge',
] as const;

const effectLstOf = (host: XmlElement): XmlElement | null => {
  for (const c of host.children) {
    if (c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'effectLst')
      return c;
  }
  return null;
};

/**
 * Puts `effect` into `host`'s effect list, replacing the one of its own kind
 * and leaving every other effect alone — a shape can carry a shadow and a glow
 * at once, and PowerPoint routinely writes both. `clearEffects` is how a caller
 * asks for the list to be emptied.
 */
const putEffect = (
  host: XmlElement,
  effect: XmlElement,
  place: EffectPlacement = effectInsertionIndex,
): void => {
  let list = effectLstOf(host);
  if (list === null) {
    list = elem(NAME_EFFECT_LST, { children: [] });
    host.children.splice(place(host), 0, list);
  }
  const rank = (name: string): number => {
    const index = (EFFECT_ORDER as readonly string[]).indexOf(name);
    // An effect the schema does not list sorts last rather than ahead of a
    // known one, so an unrecognised child cannot push a known one out of order.
    return index === -1 ? EFFECT_ORDER.length : index;
  };
  const own = rank(effect.name.localName);
  const kept = list.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        c.name.localName === effect.name.localName
      ),
  );
  let at = kept.length;
  for (let i = 0; i < kept.length; i++) {
    const c = kept[i];
    if (c?.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (rank(c.name.localName) > own) {
      at = i;
      break;
    }
  }
  kept.splice(at, 0, effect);
  list.children = kept;
};

const colorWithAlpha = (color: string, opacity: number | undefined): XmlElement => {
  const base = buildColorElement(color);
  if (opacity !== undefined && opacity >= 0 && opacity < 1) {
    const amt = Math.round(opacity * 100000);
    base.children.push(elem(NAME_ALPHA, { attrs: [attr(ATTR_VAL, String(amt))] }));
  }
  return base;
};

/**
 * Sets an outer shadow on `host`'s effect list, replacing any prior outer
 * shadow and leaving the shape's other effects in place.
 */
export const setShadow = (
  host: XmlElement,
  options: ShadowOptions = {},
  place?: EffectPlacement,
): void => {
  const color = options.color ?? '#000000';
  // blurRad and dist are ST_PositiveCoordinate (EMU, 0..27273042316900); a
  // fractional/negative/non-finite/over-max value would emit a schema-invalid
  // `<a:outerShdw>`. Validate at this boundary like every other EMU input.
  const blur = emuExtent(options.blurEmu ?? 50800, 'setShapeShadow: blurEmu');
  const dist = emuExtent(options.offsetEmu ?? 38100, 'setShapeShadow: offsetEmu');
  const angleDeg = options.angleDeg ?? 45;
  const dir = String(Math.round((((angleDeg % 360) + 360) % 360) * 60000));

  const outerShdw = elem(NAME_OUTER_SHDW, {
    attrs: [
      attr(ATTR_BLUR_RAD, String(blur)),
      attr(ATTR_DIST, String(dist)),
      attr(ATTR_DIR, dir),
      attr(ATTR_ALGN, 'tl'),
      attr(ATTR_ROT_WITH_SHAPE, '0'),
    ],
    children: [colorWithAlpha(color, options.opacity)],
  });
  putEffect(host, outerShdw, place);
};

/**
 * Sets a glow on `host`'s effect list, replacing any prior glow and leaving
 * the shape's other effects in place.
 */
export const setGlow = (host: XmlElement, options: GlowOptions, place?: EffectPlacement): void => {
  // rad is ST_PositiveCoordinate — validate like the shadow EMU inputs above.
  const rad = String(emuExtent(options.radiusEmu ?? 63500, 'setShapeGlow: radiusEmu'));
  const glow = elem(NAME_GLOW, {
    attrs: [attr(ATTR_RAD, rad)],
    children: [colorWithAlpha(options.color, options.opacity)],
  });
  putEffect(host, glow, place);
};

/** Removes any effect list from `host`. */
export const clearEffects = (host: XmlElement): void => {
  removeEffectLst(host);
};

/**
 * Removes one kind of effect, and the list with it once nothing is left — an
 * empty `<a:effectLst>` is valid but states "no effects here", which stops the
 * inheritance a run or shape without one would otherwise get.
 */
export const removeEffect = (host: XmlElement, localName: string): void => {
  const list = effectLstOf(host);
  if (list === null) return;
  list.children = list.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === localName),
  );
  if (list.children.length === 0) removeEffectLst(host);
};
