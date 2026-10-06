import { NS, attr, elem, getAttrValue, qname, type XmlElement } from '../xml/index.ts';

/** Ordered DrawingML color adjustments. Percentages are fractions; hue angles are degrees. */
export type ColorTransform =
  | { readonly kind: 'comp' | 'inv' | 'gray' | 'gamma' | 'invGamma' }
  | {
      readonly kind:
        | 'tint'
        | 'shade'
        | 'alpha'
        | 'alphaOff'
        | 'alphaMod'
        | 'hue'
        | 'hueOff'
        | 'hueMod'
        | 'sat'
        | 'satOff'
        | 'satMod'
        | 'lum'
        | 'lumOff'
        | 'lumMod'
        | 'red'
        | 'redOff'
        | 'redMod'
        | 'green'
        | 'greenOff'
        | 'greenMod'
        | 'blue'
        | 'blueOff'
        | 'blueMod';
      readonly value: number;
    };

const flags = new Set(['comp', 'inv', 'gray', 'gamma', 'invGamma']);
const values = new Set([
  'tint',
  'shade',
  'alpha',
  'alphaOff',
  'alphaMod',
  'hue',
  'hueOff',
  'hueMod',
  'sat',
  'satOff',
  'satMod',
  'lum',
  'lumOff',
  'lumMod',
  'red',
  'redOff',
  'redMod',
  'green',
  'greenOff',
  'greenMod',
  'blue',
  'blueOff',
  'blueMod',
]);
const percentageScale = 100000;
const angleScale = 60000;
const fullTurn = 360;
const minInt32 = -(2 ** 31);
const maxInt32 = 2 ** 31 - 1;

/** @internal */
export function readColorTransforms(color: XmlElement): ColorTransform[] {
  const transforms: ColorTransform[] = [];
  for (const child of color.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.dml) continue;
    const kind = child.name.localName;
    if (flags.has(kind)) {
      transforms.push({ kind: kind as Extract<ColorTransform, { value?: never }>['kind'] });
    } else if (values.has(kind)) {
      const raw = getAttrValue(child, qname('', 'val', ''));
      if (raw === null) continue;
      const value =
        Number.parseFloat(raw) /
        (kind === 'hue' || kind === 'hueOff'
          ? angleScale
          : raw.endsWith('%')
            ? 100
            : percentageScale);
      if (Number.isFinite(value))
        transforms.push({
          kind: kind as Extract<ColorTransform, { value: number }>['kind'],
          value,
        });
    }
  }
  return transforms;
}

/** @internal */
export function buildColorTransforms(transforms: readonly ColorTransform[]): XmlElement[] {
  return transforms.map((transform) => {
    if (!('value' in transform)) {
      if (!flags.has(transform.kind)) throw new RangeError('unknown color transform');
      return elem(qname('a', transform.kind, NS.dml));
    }
    const { kind, value } = transform;
    const angle = kind === 'hue' || kind === 'hueOff';
    const wire = Math.round(value * (angle ? angleScale : percentageScale));
    if (
      !values.has(kind) ||
      !Number.isFinite(value) ||
      wire < minInt32 ||
      wire > maxInt32 ||
      (['tint', 'shade', 'alpha'].includes(kind) && (value < 0 || value > 1)) ||
      (kind === 'alphaOff' && (value < -1 || value > 1)) ||
      (['alphaMod', 'hueMod'].includes(kind) && value < 0) ||
      (kind === 'hue' && (value < 0 || wire >= fullTurn * angleScale))
    )
      throw new RangeError(`invalid ${kind} color transform: ${value}`);
    return elem(qname('a', kind, NS.dml), { attrs: [attr(qname('', 'val', ''), String(wire))] });
  });
}

/** @internal */
export function colorTransformBrightness(
  transforms: readonly ColorTransform[],
): number | undefined {
  const mod = transforms.find((transform) => transform.kind === 'lumMod');
  const off = transforms.find((transform) => transform.kind === 'lumOff');
  if (!mod || !('value' in mod)) return undefined;
  const offValue = off && 'value' in off ? off.value : 0;
  if (offValue > 0 && Math.round((mod.value + offValue) * percentageScale) === percentageScale)
    return offValue;
  if (offValue === 0 && mod.value >= 0 && mod.value <= 1) return mod.value - 1;
  return undefined;
}

/** @internal */
export function colorTransformOpacity(transforms: readonly ColorTransform[]): number {
  let value = 1;
  for (const transform of transforms) {
    if (!('value' in transform)) continue;
    if (transform.kind === 'alpha') value = transform.value;
    else if (transform.kind === 'alphaMod') value *= transform.value;
    else if (transform.kind === 'alphaOff') value += transform.value;
  }
  return Math.max(0, Math.min(1, value));
}
