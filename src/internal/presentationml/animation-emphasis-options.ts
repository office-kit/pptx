// The Effect Options of the emphasis effects: how far Spin turns and which
// way, how much Grow/Shrink scales and along which axes, how transparent
// Transparency makes the shape, and the colour the colour effects change to.
//
// Unlike a fly's direction or a wheel's spokes, none of these is part of the
// preset numbers. The reference desktop app writes Spin as `presetID="8" presetSubtype="0"`
// whatever it turns through; what changes is the behaviour itself — the
// `by` of `<p:animRot>` (an `a:ST_Angle`, negative counter-clockwise), the
// `<p:by>` point of `<p:animScale>` (ECMA-376 §19.5.45, `a:ST_Percentage` per
// axis), the `style.opacity` a `<p:set>` writes, and the `<p:to>` colour of
// `<p:animClr>` (§19.5.13) or of a `<p:set>`'s `<p:clrVal>`. So they are
// written into the behaviours the preset template already holds, and read
// back from those behaviours rather than from the preset table.
//
// The reference desktop app (Mac 16.113) was captured with each effect's gallery default only
// (test/fixtures/native/animations/emphasis/). The defaults below are those
// captures; every other value is written into the same attribute the capture
// shows, which is what the schema and the captured shape settle, but has not
// been compared with a file the reference desktop app saved for that option.

import {
  type Color,
  type ColorTransform,
  asColor,
  buildColorElement,
  parseColor,
} from '../drawingml/index.ts';
import { buildColorTransforms, readColorTransforms } from '../drawingml/color-transforms.ts';
import { NS, type XmlElement, attr, firstChildElement, getAttrValue, qname } from '../xml/index.ts';

/** Which way Spin turns — the reference desktop app's Effect Options "Direction". */
export type AnimationSpinDirection = 'clockwise' | 'counterclockwise';

/**
 * Which way Grow/Shrink scales: both axes, or only the width (`'horizontal'`)
 * or only the height (`'vertical'`).
 */
export type AnimationScaleDirection = 'both' | 'horizontal' | 'vertical';

/**
 * The colour a colour emphasis effect changes to: a theme slot or an sRGB
 * value, optionally with DrawingML colour transforms — the tints and shades
 * of a theme colour palette are `lumMod` / `lumOff` on the theme slot.
 */
export type AnimationColor =
  | Color
  | {
      readonly color: Color;
      readonly colorTransforms?: readonly ColorTransform[];
    };

/** The emphasis options an effect carries, resolved: `null` where it takes none. */
export interface AnimationEmphasisOptions {
  readonly spinDirection: AnimationSpinDirection | null;
  readonly spinDegrees: number | null;
  readonly scaleDirection: AnimationScaleDirection | null;
  readonly scalePercent: number | null;
  readonly transparencyPercent: number | null;
  readonly color: AnimationColor | null;
}

export type AnimationEmphasisOptionName = keyof AnimationEmphasisOptions;

export const ANIMATION_EMPHASIS_OPTION_NAMES: readonly AnimationEmphasisOptionName[] = [
  'spinDirection',
  'spinDegrees',
  'scaleDirection',
  'scalePercent',
  'transparencyPercent',
  'color',
];

const NONE: AnimationEmphasisOptions = {
  spinDirection: null,
  spinDegrees: null,
  scaleDirection: null,
  scalePercent: null,
  transparencyPercent: null,
  color: null,
};

/** A full turn in `a:ST_Angle`, which counts sixtieth-thousandths of a degree. */
const ANGLE_PER_DEGREE = 60000;
/** `a:ST_Percentage` counts thousandths of a percent. */
const PERCENTAGE_PER_PERCENT = 1000;
const UNSCALED = 100 * PERCENTAGE_PER_PERCENT;
const INT32_MAX = 2 ** 31 - 1;
const PERCENT = 100;
// Enough places that a percentage read back from `style.opacity` is the one
// that was written, without the binary fraction showing through.
const OPACITY_DIGITS = 6;

// The captured gallery defaults: a full clockwise turn, 150 % both ways, 50 %
// transparent, Accent 2 — and Background 1 for Color Pulse.
const SPIN_DEFAULTS = { spinDirection: 'clockwise', spinDegrees: 360 } as const;
const SCALE_DEFAULTS = { scaleDirection: 'both', scalePercent: 150 } as const;
const TRANSPARENCY_DEFAULT = 50;
const COLOR_DEFAULTS: Readonly<Record<string, Color>> = {
  fillColor: 'scheme:accent2',
  fontColor: 'scheme:accent2',
  lineColor: 'scheme:accent2',
  brushColor: 'scheme:accent2',
  objectColor: 'scheme:accent2',
  growWithColor: 'scheme:accent2',
  colorPulse: 'scheme:bg1',
};

const OWNERS: Readonly<Record<AnimationEmphasisOptionName, string>> = {
  spinDirection: '"spin"',
  spinDegrees: '"spin"',
  scaleDirection: '"growShrink"',
  scalePercent: '"growShrink"',
  transparencyPercent: '"transparency"',
  color: Object.keys(COLOR_DEFAULTS)
    .map((effect) => JSON.stringify(effect))
    .join(', '),
};

const SPIN_DIRECTIONS: readonly AnimationSpinDirection[] = ['clockwise', 'counterclockwise'];
const SCALE_DIRECTIONS: readonly AnimationScaleDirection[] = ['both', 'horizontal', 'vertical'];

/** Which emphasis options `effect` takes; empty for every other effect. */
export const animationEmphasisOptionNames = (
  effect: string,
): readonly AnimationEmphasisOptionName[] => {
  if (effect === 'spin') return ['spinDirection', 'spinDegrees'];
  if (effect === 'growShrink') return ['scaleDirection', 'scalePercent'];
  if (effect === 'transparency') return ['transparencyPercent'];
  if (COLOR_DEFAULTS[effect] !== undefined) return ['color'];
  return [];
};

/** What each option is when nothing says otherwise. */
const defaultsOf = (effect: string): AnimationEmphasisOptions => {
  if (effect === 'spin') return { ...NONE, ...SPIN_DEFAULTS };
  if (effect === 'growShrink') return { ...NONE, ...SCALE_DEFAULTS };
  if (effect === 'transparency') return { ...NONE, transparencyPercent: TRANSPARENCY_DEFAULT };
  const color = COLOR_DEFAULTS[effect];
  return color === undefined ? NONE : { ...NONE, color };
};

/** The options passed for an effect, before they are checked against it. */
export type AnimationEmphasisOptionValues = Partial<Record<AnimationEmphasisOptionName, unknown>>;

const positiveNumber = (value: unknown, scale: number, label: string): number => {
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) {
    throw new RangeError(`${label} must be a positive number, not ${JSON.stringify(value)}.`);
  }
  if (Math.round(value * scale) > INT32_MAX) {
    throw new RangeError(`${label} ${value} is larger than the file can record.`);
  }
  return value;
};

const checkedColor = (value: unknown, label: string): AnimationColor => {
  const base =
    typeof value === 'object' && value !== null ? (value as { color?: unknown }).color : value;
  if (typeof base !== 'string' || parseColor(base) === null) {
    throw new RangeError(
      `${label} must be a theme colour token or "#RRGGBB", not ${JSON.stringify(base)}.`,
    );
  }
  if (typeof value === 'object' && value !== null) {
    const transforms = (value as { colorTransforms?: readonly ColorTransform[] }).colorTransforms;
    // Built once here so an invalid transform is refused before anything is written.
    if (transforms !== undefined) buildColorTransforms(transforms);
  }
  return value as AnimationColor;
};

/**
 * Validates the emphasis options passed for `effect` and fills in the defaults
 * of the ones it takes. An option the effect does not take is an error, as it
 * is for the preset options.
 */
export const resolveAnimationEmphasisOptions = (
  effect: string,
  given: AnimationEmphasisOptionValues,
  label: string,
): AnimationEmphasisOptions => {
  const takes = new Set(animationEmphasisOptionNames(effect));
  for (const name of ANIMATION_EMPHASIS_OPTION_NAMES) {
    if (given[name] !== undefined && !takes.has(name)) {
      throw new RangeError(
        `${label}: ${name} does not apply to ${JSON.stringify(effect)}; it is an Effect Option ` +
          `of ${OWNERS[name]}.`,
      );
    }
  }
  const defaults = defaultsOf(effect);
  const pick = <K extends AnimationEmphasisOptionName>(
    name: K,
    check: (value: unknown) => AnimationEmphasisOptions[K],
  ): AnimationEmphasisOptions[K] =>
    !takes.has(name) ? null : given[name] === undefined ? defaults[name] : check(given[name]);
  const among =
    <T extends string>(name: string, domain: readonly T[]) =>
    (value: unknown): T => {
      const found = domain.find((candidate) => candidate === value);
      if (found === undefined) {
        throw new RangeError(
          `${label}: ${name} ${JSON.stringify(value)} is not one of ${domain.join(', ')}.`,
        );
      }
      return found;
    };
  return {
    spinDirection: pick('spinDirection', among('spinDirection', SPIN_DIRECTIONS)),
    spinDegrees: pick('spinDegrees', (v) =>
      positiveNumber(v, ANGLE_PER_DEGREE, `${label}: spinDegrees`),
    ),
    scaleDirection: pick('scaleDirection', among('scaleDirection', SCALE_DIRECTIONS)),
    scalePercent: pick('scalePercent', (v) =>
      positiveNumber(v, PERCENTAGE_PER_PERCENT, `${label}: scalePercent`),
    ),
    transparencyPercent: pick('transparencyPercent', (v) => {
      if (typeof v !== 'number' || !Number.isFinite(v) || v < 0 || v > PERCENT) {
        throw new RangeError(
          `${label}: transparencyPercent must be a number from 0 to 100, not ${JSON.stringify(v)}.`,
        );
      }
      return v;
    }),
    color: pick('color', (v) => checkedColor(v, `${label}: color`)),
  };
};

// ---------------------------------------------------------------------------
// The behaviours, as the preset template writes them.

const isPml = (el: XmlElement, local: string): boolean =>
  el.name.namespaceURI === NS.pml && el.name.localName === local;

const NAME_CHILD_TN_LST = qname('p', 'childTnLst', NS.pml);
const NAME_C_BHVR = qname('p', 'cBhvr', NS.pml);
const NAME_ATTR_NAME_LST = qname('p', 'attrNameLst', NS.pml);
const NAME_ATTR_NAME = qname('p', 'attrName', NS.pml);
const NAME_TO = qname('p', 'to', NS.pml);
const NAME_BY = qname('p', 'by', NS.pml);
const NAME_FROM = qname('p', 'from', NS.pml);
const NAME_STR_VAL = qname('p', 'strVal', NS.pml);
const NAME_CLR_VAL = qname('p', 'clrVal', NS.pml);
const ATTR_BY = qname('', 'by', '');
const ATTR_FROM = qname('', 'from', '');
const ATTR_TO = qname('', 'to', '');
const ATTR_X = qname('', 'x', '');
const ATTR_Y = qname('', 'y', '');
const ATTR_VAL = qname('', 'val', '');
const ATTR_PR_LST = qname('', 'prLst', '');

const behavioursOf = (effectCTn: XmlElement): XmlElement[] => {
  const list = firstChildElement(effectCTn, NAME_CHILD_TN_LST);
  return list === null
    ? []
    : list.children.filter(
        (c): c is XmlElement => c.kind === 'element' && c.name.namespaceURI === NS.pml,
      );
};

const attrNameOf = (behaviour: XmlElement): string | null => {
  const cBhvr = firstChildElement(behaviour, NAME_C_BHVR);
  const list = cBhvr === null ? null : firstChildElement(cBhvr, NAME_ATTR_NAME_LST);
  const name = list === null ? null : firstChildElement(list, NAME_ATTR_NAME);
  return name?.children.find((c) => c.kind === 'text')?.data ?? null;
};

const setAttr = (el: XmlElement, name: string, value: string): void => {
  const has = el.attrs.some((a) => a.name.namespaceURI === '' && a.name.localName === name);
  el.attrs = has
    ? el.attrs.map((a) =>
        a.name.namespaceURI === '' && a.name.localName === name ? { ...a, value } : a,
      )
    : [...el.attrs, attr(qname('', name, ''), value)];
};

/** The element a colour behaviour states its colour in: `<p:to>`, or the `<p:clrVal>` in it. */
const colorHosts = (effectCTn: XmlElement): XmlElement[] => {
  const out: XmlElement[] = [];
  for (const behaviour of behavioursOf(effectCTn)) {
    const to = firstChildElement(behaviour, NAME_TO);
    if (to === null) continue;
    if (isPml(behaviour, 'animClr')) out.push(to);
    else if (isPml(behaviour, 'set')) {
      const clrVal = firstChildElement(to, NAME_CLR_VAL);
      if (clrVal !== null) out.push(clrVal);
    }
  }
  return out;
};

const colorElement = (value: AnimationColor): XmlElement => {
  const base = typeof value === 'string' ? value : value.color;
  const el = buildColorElement(base);
  if (typeof value !== 'string' && value.colorTransforms !== undefined) {
    el.children = buildColorTransforms(value.colorTransforms);
  }
  return el;
};

const opacityText = (transparencyPercent: number): string =>
  String(Number((1 - transparencyPercent / PERCENT).toFixed(OPACITY_DIGITS)));

const percentText = (percent: number): string =>
  String(Math.round(percent * PERCENTAGE_PER_PERCENT));

/**
 * Writes the emphasis options into the behaviours of a freshly built effect,
 * in place. The template already holds the gallery default, so writing the
 * defaults changes nothing.
 */
export const applyAnimationEmphasisOptions = (
  effectCTn: XmlElement,
  options: AnimationEmphasisOptions,
): void => {
  for (const behaviour of behavioursOf(effectCTn)) {
    if (isPml(behaviour, 'animRot') && options.spinDegrees !== null) {
      const sign = options.spinDirection === 'counterclockwise' ? -1 : 1;
      setAttr(behaviour, 'by', String(sign * Math.round(options.spinDegrees * ANGLE_PER_DEGREE)));
    }
    if (isPml(behaviour, 'animScale') && options.scalePercent !== null) {
      const by = firstChildElement(behaviour, NAME_BY);
      if (by === null) continue;
      const scaled = percentText(options.scalePercent);
      const unscaled = String(UNSCALED);
      setAttr(by, 'x', options.scaleDirection === 'vertical' ? unscaled : scaled);
      setAttr(by, 'y', options.scaleDirection === 'horizontal' ? unscaled : scaled);
    }
    if (options.transparencyPercent !== null) {
      const opacity = opacityText(options.transparencyPercent);
      if (isPml(behaviour, 'set') && attrNameOf(behaviour) === 'style.opacity') {
        const strVal = firstChildElement(firstChildElement(behaviour, NAME_TO)!, NAME_STR_VAL);
        if (strVal !== null) setAttr(strVal, 'val', opacity);
      }
      // The reference desktop app repeats the opacity as the image filter's property list.
      if (isPml(behaviour, 'animEffect') && getAttrValue(behaviour, ATTR_PR_LST) !== null) {
        setAttr(behaviour, 'prLst', `opacity: ${opacity}`);
      }
    }
  }
  if (options.color !== null) {
    for (const host of colorHosts(effectCTn)) host.children = [colorElement(options.color)];
  }
};

const strictInt = (raw: string | null): number | null => {
  if (raw === null || !/^-?\d+$/.test(raw)) return null;
  const n = Number(raw);
  return Number.isSafeInteger(n) ? n : null;
};

/** The colour a `<p:to>` / `<p:clrVal>` states, when it is one this library can write again. */
const readColor = (host: XmlElement): AnimationColor | null => {
  const elements = host.children.filter((c): c is XmlElement => c.kind === 'element');
  const el = elements[0];
  if (elements.length !== 1 || el === undefined || el.name.namespaceURI !== NS.dml) return null;
  const val = getAttrValue(el, ATTR_VAL);
  if (val === null) return null;
  const base =
    el.name.localName === 'srgbClr'
      ? asColor(`#${val.toUpperCase()}`)
      : el.name.localName === 'schemeClr'
        ? asColor(`scheme:${val}`)
        : null;
  if (base === null) return null;
  const transforms = readColorTransforms(el);
  // A child the transform reader skipped is something the writer would drop.
  if (transforms.length !== el.children.filter((c) => c.kind === 'element').length) return null;
  return transforms.length === 0 ? base : { color: base, colorTransforms: transforms };
};

const sameColor = (a: AnimationColor, b: AnimationColor): boolean =>
  JSON.stringify(a) === JSON.stringify(b);

const withOptions = (values: Partial<AnimationEmphasisOptions>): AnimationEmphasisOptions => ({
  ...NONE,
  ...values,
});

/**
 * The emphasis options an effect node states, or `null` when its behaviours
 * say something these options cannot — a turn fixed by `from` / `to`, a scale
 * that stretches both axes by different amounts, two colours — so the effect
 * is one this library reads rather than names. An effect that takes none of
 * these options states none.
 */
export const readAnimationEmphasisOptions = (
  effect: string,
  effectCTn: XmlElement,
): AnimationEmphasisOptions | null => {
  const names = animationEmphasisOptionNames(effect);
  if (names.length === 0) return NONE;
  const behaviours = behavioursOf(effectCTn);
  if (effect === 'spin') {
    const rotation = behaviours[0];
    if (behaviours.length !== 1 || rotation === undefined || !isPml(rotation, 'animRot')) {
      return null;
    }
    if (getAttrValue(rotation, ATTR_FROM) !== null || getAttrValue(rotation, ATTR_TO) !== null) {
      return null;
    }
    const by = strictInt(getAttrValue(rotation, ATTR_BY));
    if (by === null || by === 0) return null;
    return withOptions({
      spinDirection: by < 0 ? 'counterclockwise' : 'clockwise',
      spinDegrees: Math.abs(by) / ANGLE_PER_DEGREE,
    });
  }
  if (effect === 'growShrink') {
    const scale = behaviours[0];
    if (behaviours.length !== 1 || scale === undefined || !isPml(scale, 'animScale')) return null;
    if (
      firstChildElement(scale, NAME_FROM) !== null ||
      firstChildElement(scale, NAME_TO) !== null
    ) {
      return null;
    }
    const by = firstChildElement(scale, NAME_BY);
    const x = by === null ? null : strictInt(getAttrValue(by, ATTR_X));
    const y = by === null ? null : strictInt(getAttrValue(by, ATTR_Y));
    if (x === null || y === null || x <= 0 || y <= 0) return null;
    const scaleDirection: AnimationScaleDirection | null =
      x === y ? 'both' : y === UNSCALED ? 'horizontal' : x === UNSCALED ? 'vertical' : null;
    if (scaleDirection === null) return null;
    const percent = (scaleDirection === 'vertical' ? y : x) / PERCENTAGE_PER_PERCENT;
    return withOptions({ scaleDirection, scalePercent: percent });
  }
  if (effect === 'transparency') {
    const set = behaviours.find((b) => isPml(b, 'set') && attrNameOf(b) === 'style.opacity');
    const to = set === undefined ? null : firstChildElement(set, NAME_TO);
    const strVal = to === null ? null : firstChildElement(to, NAME_STR_VAL);
    const raw = strVal === null ? null : getAttrValue(strVal, ATTR_VAL);
    const opacity = raw === null || !/^\d*\.?\d+$/.test(raw.trim()) ? Number.NaN : Number(raw);
    if (!(opacity >= 0 && opacity <= 1)) return null;
    return withOptions({
      transparencyPercent: Number(((1 - opacity) * PERCENT).toFixed(OPACITY_DIGITS - 2)),
    });
  }
  const hosts = colorHosts(effectCTn);
  const colors = hosts.map(readColor);
  const first = colors[0];
  if (first === undefined || first === null) return null;
  if (!colors.every((c) => c !== null && sameColor(c, first))) return null;
  return withOptions({ color: first });
};
