// Solid-fill mutation for shapes and slide backgrounds.
//
// The spec places the fill element after geometry (`prstGeom`/`custGeom`)
// and before the line (`a:ln`) on a shape's `<p:spPr>`. We re-insert
// rather than append so re-applying a fill stays clean across calls.
//
// `setSolidFill(host, color)` accepts the host element that wraps the
// fill (`p:spPr` for shapes, `p:bgPr` for backgrounds). It removes any
// previous fill choice (`noFill`/`solidFill`/`gradFill`/`blipFill`/
// `pattFill`/`grpFill`) before inserting the new `solidFill`.

import { NS, type XmlElement, attr, elem, qname } from '../xml/index.ts';
import { buildColorElement } from './color.ts';

const NAME_SOLID_FILL = qname('a', 'solidFill', NS.dml);
const NAME_NO_FILL = qname('a', 'noFill', NS.dml);
const NAME_GRAD_FILL = qname('a', 'gradFill', NS.dml);
const NAME_GS_LST = qname('a', 'gsLst', NS.dml);
const NAME_GS = qname('a', 'gs', NS.dml);
const NAME_LIN = qname('a', 'lin', NS.dml);
const ATTR_POS = qname('', 'pos', '');
const ATTR_ANG = qname('', 'ang', '');
const ATTR_SCALED = qname('', 'scaled', '');
const ATTR_FLIP = qname('', 'flip', '');
const ATTR_ROT_WITH_SHAPE = qname('', 'rotWithShape', '');

const FILL_CHOICE_LOCAL_NAMES = new Set([
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
]);

const removeAnyFill = (host: XmlElement): void => {
  host.children = host.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        FILL_CHOICE_LOCAL_NAMES.has(c.name.localName)
      ),
  );
};

/**
 * Returns the index where the fill should be inserted on `host` per the
 * schema's child-element sequence. For `<p:spPr>` that's right after the
 * geometry element; for `<p:bgPr>` it's the start.
 *
 * The schema is more nuanced (xfrm → geometry → fill → ln → effects →
 * scene3d → sp3d → extLst) but PowerPoint tolerates any of those slots
 * being absent, and we only care about staying ahead of `a:ln`.
 */
const fillInsertionIndex = (host: XmlElement): number => {
  for (let i = 0; i < host.children.length; i++) {
    const c = host.children[i];
    if (c?.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (c.name.localName === 'ln') return i;
    if (c.name.localName === 'effectLst' || c.name.localName === 'effectDag') return i;
    if (c.name.localName === 'scene3d' || c.name.localName === 'sp3d') return i;
    if (c.name.localName === 'extLst') return i;
  }
  return host.children.length;
};

/** Sets `<a:solidFill>` on `host`, removing any previous fill choice. */
export const setSolidFill = (host: XmlElement, color: string): void => {
  removeAnyFill(host);
  const fill = elem(NAME_SOLID_FILL, { children: [buildColorElement(color)] });
  host.children.splice(fillInsertionIndex(host), 0, fill);
};

/** Sets `<a:noFill>` on `host`, removing any previous fill choice. */
export const setNoFill = (host: XmlElement): void => {
  removeAnyFill(host);
  host.children.splice(fillInsertionIndex(host), 0, elem(NAME_NO_FILL));
};

/**
 * Removes any fill choice from `host` entirely. The shape then inherits
 * its fill from the layout / master placeholder it descends from.
 */
export const clearFill = (host: XmlElement): void => {
  removeAnyFill(host);
};

/** One stop in a linear gradient. */
export interface GradientStop {
  /** Position on the gradient axis, 0-1. */
  readonly offset: number;
  /** `#RRGGBB`, bare `RRGGBB`, or a scheme color token. */
  readonly color: string;
}

export interface GradientFillOptions {
  /** Two or more color stops along the axis. */
  readonly stops: ReadonlyArray<GradientStop>;
  /**
   * Gradient axis angle, in degrees. ECMA-376 measures clockwise from
   * the right (3 o'clock); `0` is a left-to-right gradient, `90` is
   * top-to-bottom, `180` right-to-left, `270` bottom-to-top. Defaults
   * to `90` (top → bottom).
   */
  readonly angleDeg?: number;
}

/** One of ECMA-376's `ST_PresetPatternVal` tokens (`pct50`, `dkUpDiag`, ...). */
export type PatternPreset = string;

export interface PatternFillOptions {
  /** Preset pattern token, e.g. `'pct50'`, `'dkUpDiag'`, `'wave'`. */
  readonly preset: PatternPreset;
  /** Foreground (pattern stroke) color. `#RRGGBB`, bare `RRGGBB`, or scheme token. */
  readonly foreground: string;
  /** Background (fill behind the pattern) color. */
  readonly background: string;
}

const NAME_PATT_FILL = qname('a', 'pattFill', NS.dml);
const NAME_FG_CLR = qname('a', 'fgClr', NS.dml);
const NAME_BG_CLR = qname('a', 'bgClr', NS.dml);
const ATTR_PRST = qname('', 'prst', '');

/**
 * Sets `<a:pattFill>` on `host` with the given preset + colors.
 * Replaces any previous fill choice.
 */
export const setPatternFill = (host: XmlElement, options: PatternFillOptions): void => {
  removeAnyFill(host);
  const pattFill = elem(NAME_PATT_FILL, {
    attrs: [attr(ATTR_PRST, options.preset)],
    children: [
      elem(NAME_FG_CLR, { children: [buildColorElement(options.foreground)] }),
      elem(NAME_BG_CLR, { children: [buildColorElement(options.background)] }),
    ],
  });
  host.children.splice(fillInsertionIndex(host), 0, pattFill);
};

/** Sets `<a:gradFill>` on `host`, replacing any previous fill choice. */
export const setGradientFill = (host: XmlElement, options: GradientFillOptions): void => {
  if (options.stops.length < 2) {
    throw new Error('gradient fill requires at least two stops');
  }
  removeAnyFill(host);

  const stops = options.stops.map((s) => {
    if (!Number.isFinite(s.offset) || s.offset < 0 || s.offset > 1) {
      throw new RangeError(`gradient stop offset must be in [0, 1], got ${s.offset}`);
    }
    const posST = String(Math.round(s.offset * 100000));
    return elem(NAME_GS, {
      attrs: [attr(ATTR_POS, posST)],
      children: [buildColorElement(s.color)],
    });
  });

  const angleDeg = options.angleDeg ?? 90;
  // ECMA-376 ST_PositiveFixedAngle: 60000 units per degree, range
  // [0, 21600000). Normalize negatives via modulo.
  const norm = ((angleDeg % 360) + 360) % 360;
  const angleAttr = String(Math.round(norm * 60000));

  const grad = elem(NAME_GRAD_FILL, {
    attrs: [attr(ATTR_FLIP, 'none'), attr(ATTR_ROT_WITH_SHAPE, '1')],
    children: [
      elem(NAME_GS_LST, { children: stops }),
      elem(NAME_LIN, { attrs: [attr(ATTR_ANG, angleAttr), attr(ATTR_SCALED, '0')] }),
    ],
  });
  host.children.splice(fillInsertionIndex(host), 0, grad);
};
