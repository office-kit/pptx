// Shape effects — `<a:effectLst>` builders.
//
// Covers the most-used PowerPoint effects: outer shadow + glow. The
// element ordering on `<p:spPr>` is `xfrm → geometry → fill → ln →
// effectLst → scene3d → sp3d → extLst`. Callers locate the right
// insertion slot using `effectInsertionIndex`.

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
}

/**
 * Computes the index inside `host.children` where an `<a:effectLst>`
 * should be inserted to satisfy the spec's child ordering on
 * `<p:spPr>`.
 */
const effectInsertionIndex = (host: XmlElement): number => {
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
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        c.name.localName === 'effectLst'
      ),
  );
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
 * Sets an outer shadow on `host`'s effect list. Replaces any prior
 * `<a:effectLst>` entirely (we treat shadow + glow as mutually
 * exclusive in v1 — multi-effect stacks are a post-1.0 enhancement).
 */
export const setShadow = (host: XmlElement, options: ShadowOptions = {}): void => {
  removeEffectLst(host);
  const color = options.color ?? '#000000';
  const blur = options.blurEmu ?? 50800;
  const dist = options.offsetEmu ?? 38100;
  const angleDeg = options.angleDeg ?? 45;
  const dir = String(Math.round(((angleDeg % 360 + 360) % 360) * 60000));

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
  const effectLst = elem(NAME_EFFECT_LST, { children: [outerShdw] });
  host.children.splice(effectInsertionIndex(host), 0, effectLst);
};

/**
 * Sets a glow on `host`'s effect list. Replaces any prior
 * `<a:effectLst>`.
 */
export const setGlow = (host: XmlElement, options: GlowOptions): void => {
  removeEffectLst(host);
  const rad = String(options.radiusEmu ?? 63500);
  const glow = elem(NAME_GLOW, {
    attrs: [attr(ATTR_RAD, rad)],
    children: [buildColorElement(options.color)],
  });
  const effectLst = elem(NAME_EFFECT_LST, { children: [glow] });
  host.children.splice(effectInsertionIndex(host), 0, effectLst);
};

/** Removes any effect list from `host`. */
export const clearEffects = (host: XmlElement): void => {
  removeEffectLst(host);
};
