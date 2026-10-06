// Shape effects — `<a:effectLst>` builders.
//
// Covers the most-used PowerPoint effects: outer shadow, glow, and reflection. The
// element ordering on `<p:spPr>` is `xfrm → geometry → fill → ln →
// effectLst → scene3d → sp3d → extLst`. Callers locate the right
// insertion slot using `effectInsertionIndex`.

import type { Color } from './color.ts';
import { emuExtent } from '../bounds.ts';
import { NS, type XmlElement, attr, elem, qname } from '../xml/index.ts';
import { buildColorElement } from './color.ts';
import {
  type ColorTransform,
  buildColorTransforms,
  colorTransformOpacity,
} from './color-transforms.ts';

const NAME_EFFECT_LST = qname('a', 'effectLst', NS.dml);
const NAME_OUTER_SHDW = qname('a', 'outerShdw', NS.dml);
const NAME_INNER_SHDW = qname('a', 'innerShdw', NS.dml);
const NAME_GLOW = qname('a', 'glow', NS.dml);
const NAME_REFLECTION = qname('a', 'reflection', NS.dml);
const NAME_ALPHA = qname('a', 'alpha', NS.dml);

const ATTR_BLUR_RAD = qname('', 'blurRad', '');
const ATTR_DIST = qname('', 'dist', '');
const ATTR_DIR = qname('', 'dir', '');
const ATTR_ALGN = qname('', 'algn', '');
const ATTR_ROT_WITH_SHAPE = qname('', 'rotWithShape', '');
const ATTR_RAD = qname('', 'rad', '');
const ATTR_ST_A = qname('', 'stA', '');
const ATTR_ST_POS = qname('', 'stPos', '');
const ATTR_END_A = qname('', 'endA', '');
const ATTR_END_POS = qname('', 'endPos', '');
const ATTR_FADE_DIR = qname('', 'fadeDir', '');
const ATTR_SX = qname('', 'sx', '');
const ATTR_SY = qname('', 'sy', '');
const ATTR_KX = qname('', 'kx', '');
const ATTR_KY = qname('', 'ky', '');
const ATTR_VAL = qname('', 'val', '');

const PERCENTAGE_UNITS = 100000;
const ALPHA_TRANSFORMS = new Set(['alpha', 'alphaMod', 'alphaOff']);
const ANGLE_UNITS_PER_DEGREE = 60000;
const FULL_TURN_DEGREES = 360;
const FULL_TURN_UNITS = FULL_TURN_DEGREES * ANGLE_UNITS_PER_DEGREE;
const RIGHT_ANGLE_UNITS = 90 * ANGLE_UNITS_PER_DEGREE;
const SIGNED_INT_MIN = -2147483648;
const SIGNED_INT_MAX = 2147483647;

export interface ShadowOptions {
  /** `#RRGGBB`, bare `RRGGBB`, or scheme token. Defaults to black. */
  readonly color?: Color;
  /**
   * Ordered adjustments to `color` (`<a:lumMod>`, `<a:tint>`, ...), written
   * as its children — the same field gradient stops take. Requires `color`.
   * An `opacity` that differs from the one these transforms state wins.
   */
  readonly colorTransforms?: readonly ColorTransform[];
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
  /** Transform anchor; defaults to top left (`tl`) for authored shadows. */
  readonly alignment?: 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br';
  /** Whether the shadow rotates with its shape; defaults to false. */
  readonly rotateWithShape?: boolean;
}

/** Character or shape inner-shadow parameters from `<a:innerShdw>`. */
export interface InnerShadowOptions {
  /** `#RRGGBB`, bare `RRGGBB`, or scheme token. Defaults to black. */
  readonly color?: Color;
  /**
   * Ordered adjustments to `color` (`<a:lumMod>`, `<a:tint>`, ...), written
   * as its children — the same field gradient stops take. Requires `color`.
   * An `opacity` that differs from the one these transforms state wins.
   */
  readonly colorTransforms?: readonly ColorTransform[];
  /** Edge blur in EMU. Defaults to 50800 (4pt). */
  readonly blurEmu?: number;
  /** Offset distance in EMU. Defaults to 38100 (3pt). */
  readonly offsetEmu?: number;
  /** Direction in degrees, measured clockwise from the right. Defaults to 45°. */
  readonly angleDeg?: number;
  /** Opacity (0–1). Defaults to fully opaque. */
  readonly opacity?: number;
}

export interface GlowOptions {
  /** `#RRGGBB`, bare `RRGGBB`, or scheme token. */
  readonly color: Color;
  /**
   * Ordered adjustments to `color` (`<a:lumMod>`, `<a:tint>`, ...), written
   * as its children — the same field gradient stops take. Requires `color`.
   * An `opacity` that differs from the one these transforms state wins.
   */
  readonly colorTransforms?: readonly ColorTransform[];
  /** Glow radius in EMU. Defaults to 63500 (5pt). */
  readonly radiusEmu?: number;
  /** Opacity (0–1). Defaults to fully opaque. */
  readonly opacity?: number;
}

/** Character reflection parameters from DrawingML's `<a:reflection>`. */
export interface ReflectionOptions {
  /** Blur radius in EMU; defaults to zero. */
  readonly blurEmu?: number;
  /** Offset distance in EMU; defaults to zero. */
  readonly offsetEmu?: number;
  /** Offset direction in clockwise degrees from right; defaults to zero. */
  readonly angleDeg?: number;
  /** Ending opacity (0–1); defaults to zero (`endA`). */
  readonly opacity?: number;
  /** Starting opacity (0–1); defaults to one. */
  readonly startOpacity?: number;
  /** Starting fade position (0–1); defaults to zero. */
  readonly startPosition?: number;
  /** Ending fade position (0–1); defaults to one. */
  readonly endPosition?: number;
  /** Fade direction in clockwise degrees from right; defaults to 90. */
  readonly fadeDirection?: number;
  /** Horizontal scale; defaults to one. Negative values mirror horizontally. */
  readonly scaleX?: number;
  /** Vertical scale; defaults to one. Negative values mirror vertically. */
  readonly scaleY?: number;
  /** Horizontal skew in degrees, strictly between -90 and 90. */
  readonly skewX?: number;
  /** Vertical skew in degrees, strictly between -90 and 90. */
  readonly skewY?: number;
  /** Transform anchor; defaults to bottom center (`b`). */
  readonly alignment?: 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br';
  /** Whether the effect rotates with its shape; defaults to true. */
  readonly rotateWithShape?: boolean;
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

// The effect's color element: the transforms in the caller's order, then the
// opacity as a trailing `<a:alpha>` — the order PowerPoint writes
// (`<a:lumMod/><a:alpha/>`). An opacity the transforms already state is not
// repeated, so a color read back from a deck writes back unchanged.
const effectColor = (
  options: {
    readonly color?: Color;
    readonly colorTransforms?: readonly ColorTransform[];
    readonly opacity?: number;
  },
  caller: string,
): XmlElement => {
  const { color, colorTransforms, opacity } = options;
  if (colorTransforms !== undefined && color === undefined)
    throw new Error(`${caller}: colorTransforms requires color`);
  const base = buildColorElement(color ?? '#000000');
  const transforms = colorTransforms ?? [];
  base.children = buildColorTransforms(transforms);
  if (
    opacity !== undefined &&
    (colorTransforms === undefined || opacity !== colorTransformOpacity(transforms))
  ) {
    base.children = base.children.filter(
      (c) => !(c.kind === 'element' && ALPHA_TRANSFORMS.has(c.name.localName)),
    );
    if (opacity >= 0 && opacity < 1)
      base.children.push(
        elem(NAME_ALPHA, {
          attrs: [attr(ATTR_VAL, String(Math.round(opacity * PERCENTAGE_UNITS)))],
        }),
      );
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
      attr(ATTR_ALGN, options.alignment ?? 'tl'),
      attr(ATTR_ROT_WITH_SHAPE, options.rotateWithShape === true ? '1' : '0'),
    ],
    children: [effectColor(options, 'setShapeShadow')],
  });
  putEffect(host, outerShdw, place);
};

/** Sets an inner shadow, replacing only a prior inner shadow. */
export const setInnerShadow = (
  host: XmlElement,
  options: InnerShadowOptions = {},
  place?: EffectPlacement,
): void => {
  const blur = emuExtent(options.blurEmu ?? 50800, 'setInnerShadow: blurEmu');
  const dist = emuExtent(options.offsetEmu ?? 38100, 'setInnerShadow: offsetEmu');
  const angleDeg = options.angleDeg ?? 45;
  if (!Number.isFinite(angleDeg)) throw new RangeError('setInnerShadow: angleDeg must be finite');
  const opacity = options.opacity;
  if (opacity !== undefined && (!Number.isFinite(opacity) || opacity < 0 || opacity > 1))
    throw new RangeError('setInnerShadow: opacity must be in [0, 1]');
  const dir = String(
    Math.round(
      (((angleDeg % FULL_TURN_DEGREES) + FULL_TURN_DEGREES) % FULL_TURN_DEGREES) *
        ANGLE_UNITS_PER_DEGREE,
    ) % FULL_TURN_UNITS,
  );
  const innerShdw = elem(NAME_INNER_SHDW, {
    attrs: [attr(ATTR_BLUR_RAD, String(blur)), attr(ATTR_DIST, String(dist)), attr(ATTR_DIR, dir)],
    children: [effectColor(options, 'setInnerShadow')],
  });
  putEffect(host, innerShdw, place);
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
    children: [effectColor(options, 'setShapeGlow')],
  });
  putEffect(host, glow, place);
};

/** Sets a reflection on `host`'s effect list, preserving other effects. */
export const setReflection = (
  host: XmlElement,
  options: ReflectionOptions = {},
  place?: EffectPlacement,
): void => {
  const blur = emuExtent(options.blurEmu ?? 0, 'setReflection: blurEmu');
  const dist = emuExtent(options.offsetEmu ?? 0, 'setReflection: offsetEmu');
  const angleDeg = options.angleDeg ?? 0;
  if (!Number.isFinite(angleDeg)) throw new RangeError('setReflection: angleDeg must be finite');
  const fraction = (value: number | undefined, field: string): string | undefined => {
    if (value === undefined) return undefined;
    if (!Number.isFinite(value) || value < 0 || value > 1)
      throw new RangeError(`${field} must be in [0, 1]`);
    return String(Math.round(value * PERCENTAGE_UNITS));
  };
  const scalePercentage = (value: number | undefined, field: string): string | undefined => {
    if (value === undefined) return undefined;
    if (!Number.isFinite(value)) throw new RangeError(`${field} must be finite`);
    const units = Math.round(value * PERCENTAGE_UNITS);
    if (!Number.isSafeInteger(units) || units < SIGNED_INT_MIN || units > SIGNED_INT_MAX)
      throw new RangeError(`${field} must fit an OOXML percentage integer`);
    return String(units);
  };
  const fixedAngle = (value: number | undefined, field: string): string | undefined => {
    if (value === undefined) return undefined;
    if (!Number.isFinite(value)) throw new RangeError(`${field} must be finite`);
    const units = Math.round(value * ANGLE_UNITS_PER_DEGREE);
    if (units <= -RIGHT_ANGLE_UNITS || units >= RIGHT_ANGLE_UNITS)
      throw new RangeError(`${field} must be strictly between -90 and 90 degrees`);
    return String(units);
  };
  const direction =
    Math.round(
      (((angleDeg % FULL_TURN_DEGREES) + FULL_TURN_DEGREES) % FULL_TURN_DEGREES) *
        ANGLE_UNITS_PER_DEGREE,
    ) % FULL_TURN_UNITS;
  const positiveAngle = (value: number | undefined, field: string): string | undefined => {
    if (value === undefined) return undefined;
    if (!Number.isFinite(value)) throw new RangeError(`${field} must be finite`);
    return String(
      Math.round(
        (((value % FULL_TURN_DEGREES) + FULL_TURN_DEGREES) % FULL_TURN_DEGREES) *
          ANGLE_UNITS_PER_DEGREE,
      ) % FULL_TURN_UNITS,
    );
  };
  const optionalAttr = (name: ReturnType<typeof qname>, value: string | undefined) =>
    value === undefined ? [] : [attr(name, value)];
  const attrs = [
    attr(ATTR_BLUR_RAD, String(blur)),
    attr(ATTR_DIST, String(dist)),
    attr(ATTR_DIR, String(direction)),
    attr(ATTR_ALGN, options.alignment ?? 'b'),
    attr(ATTR_ROT_WITH_SHAPE, options.rotateWithShape === false ? '0' : '1'),
    ...optionalAttr(ATTR_ST_A, fraction(options.startOpacity, 'setReflection: startOpacity')),
    ...optionalAttr(ATTR_ST_POS, fraction(options.startPosition, 'setReflection: startPosition')),
    ...optionalAttr(ATTR_END_A, fraction(options.opacity, 'setReflection: opacity')),
    ...optionalAttr(ATTR_END_POS, fraction(options.endPosition, 'setReflection: endPosition')),
    ...optionalAttr(
      ATTR_FADE_DIR,
      positiveAngle(options.fadeDirection, 'setReflection: fadeDirection'),
    ),
    ...optionalAttr(ATTR_SX, scalePercentage(options.scaleX, 'setReflection: scaleX')),
    ...optionalAttr(ATTR_SY, scalePercentage(options.scaleY, 'setReflection: scaleY')),
    ...optionalAttr(ATTR_KX, fixedAngle(options.skewX, 'setReflection: skewX')),
    ...optionalAttr(ATTR_KY, fixedAngle(options.skewY, 'setReflection: skewY')),
  ];
  putEffect(host, elem(NAME_REFLECTION, { attrs }), place);
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
