// Builds the `<p:transition>` element that controls how PowerPoint
// animates from this slide to the next.
//
// Per ECMA-376 Part 1 §19.5.51 the transition carries:
//   - `spd` attribute: `slow` | `med` | `fast` (default `med`)
//   - `advClick` attribute: `1` to advance on click (default), `0` to disable
//   - `advTm` attribute: auto-advance time in milliseconds
//   - Exactly one child element from the effect catalog (fade, push,
//     cover, wipe, split, cut, dissolve, checker, blinds, randomBar,
//     zoom, circle, diamond, plus, wedge, ...).
//
// Some effects accept a direction attribute (`dir`); the schema constrains
// the legal values per element. We don't enforce that — let the user pass
// what they want; xmllint will catch invalid combos in the test layer.

import { type XmlElement, NS, attr, elem, qname } from '../xml/index.ts';

const NAME_TRANSITION = qname('p', 'transition', NS.pml);
const ATTR_SPD = qname('', 'spd', '');
const ATTR_ADV_CLICK = qname('', 'advClick', '');
const ATTR_ADV_TM = qname('', 'advTm', '');
const ATTR_DIR = qname('', 'dir', '');
const ATTR_ORIENT = qname('', 'orient', '');
const ATTR_THRU_BLK = qname('', 'thruBlk', '');

/**
 * Transition effect token. Maps to a `<p:{token}/>` child of
 * `<p:transition>`. The list covers the effects all current PowerPoint
 * versions emit; pass any other ECMA-376-permitted local name as a raw
 * string for forward compatibility.
 */
export type TransitionEffect =
  | 'none'
  | 'fade'
  | 'push'
  | 'cover'
  | 'wipe'
  | 'split'
  | 'cut'
  | 'dissolve'
  | 'checker'
  | 'blinds'
  | 'randomBar'
  | 'zoom'
  | 'circle'
  | 'diamond'
  | 'plus'
  | 'wedge'
  | 'newsflash';

export interface TransitionOptions {
  effect: TransitionEffect | string;
  /** Effect speed. Defaults to omitted (PowerPoint treats absence as `med`). */
  speed?: 'slow' | 'med' | 'fast';
  /** Direction hint accepted by some effects (`l`/`r`/`u`/`d`/`in`/`out`/...). */
  direction?: string;
  /** For `split`: orientation token (`horz` / `vert`). */
  orientation?: 'horz' | 'vert';
  /** For `fade`: pass `true` to fade through black. */
  thruBlack?: boolean;
  /** Whether clicking advances; default `true` (PowerPoint's default). */
  advanceOnClick?: boolean;
  /**
   * Milliseconds to auto-advance after this slide. Omit for click-only
   * advance.
   */
  advanceAfterMs?: number;
}

const buildEffectElement = (opts: TransitionOptions): XmlElement => {
  const name = qname('p', opts.effect, NS.pml);
  const attrs = [];
  if (opts.direction !== undefined) attrs.push(attr(ATTR_DIR, opts.direction));
  if (opts.orientation !== undefined) attrs.push(attr(ATTR_ORIENT, opts.orientation));
  if (opts.thruBlack) attrs.push(attr(ATTR_THRU_BLK, '1'));
  return elem(name, { attrs });
};

/** Returns a complete `<p:transition>` element. */
export const buildTransition = (opts: TransitionOptions): XmlElement => {
  const attrs = [];
  if (opts.speed !== undefined) attrs.push(attr(ATTR_SPD, opts.speed));
  if (opts.advanceOnClick === false) attrs.push(attr(ATTR_ADV_CLICK, '0'));
  if (opts.advanceAfterMs !== undefined) {
    attrs.push(attr(ATTR_ADV_TM, String(Math.round(opts.advanceAfterMs))));
  }
  return elem(NAME_TRANSITION, {
    attrs,
    children: [buildEffectElement(opts)],
  });
};
