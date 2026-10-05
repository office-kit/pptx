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
// The effect attributes are not interchangeable: each effect element maps to a
// distinct CT type that only permits certain attributes (CT_OptionalBlackTransition
// → thruBlk; CT_SplitTransition → orient + dir; the direction families → dir).
// Emitting an attribute on an effect that doesn't allow it is schema-invalid, so
// buildEffectElement gates each attribute by the effect that accepts it.

import { boundedInt, oneOf, unsignedIntMs } from '../bounds.ts';
import {
  type XmlAttr,
  type XmlElement,
  NS,
  attr,
  cloneElement,
  elem,
  qname,
} from '../xml/index.ts';

const NAME_TRANSITION = qname('p', 'transition', NS.pml);
const ATTR_SPD = qname('', 'spd', '');
const ATTR_ADV_CLICK = qname('', 'advClick', '');
const ATTR_ADV_TM = qname('', 'advTm', '');
const ATTR_DIR = qname('', 'dir', '');
const ATTR_ORIENT = qname('', 'orient', '');
const ATTR_THRU_BLK = qname('', 'thruBlk', '');
const ATTR_P14_DUR = qname('p14', 'dur', NS.p14);
const ATTR_REQUIRES = qname('', 'Requires', '');
const NAME_ALTERNATE_CONTENT = qname('mc', 'AlternateContent', NS.mc);
const NAME_CHOICE = qname('mc', 'Choice', NS.mc);
const NAME_FALLBACK = qname('mc', 'Fallback', NS.mc);

// PowerPoint's own durations for the three ECMA-376 speeds.
const FAST_MS = 500;
const MED_MS = 750;
const speedForDuration = (ms: number): 'slow' | 'med' | 'fast' =>
  ms <= FAST_MS ? 'fast' : ms <= MED_MS ? 'med' : 'slow';

/**
 * Every transition effect element name in `CT_SlideTransition`'s choice
 * (ECMA-376 pml.xsd). The token is emitted verbatim as `<p:{token}/>`, so this
 * list is both the type's domain and the write-time validation domain —
 * keeping them one declaration stops them drifting apart.
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
  | 'newsflash'
  | 'wheel';

export interface TransitionOptions {
  effect: TransitionEffect;
  /** Effect speed. Defaults to omitted (PowerPoint treats absence as `med`). */
  speed?: 'slow' | 'med' | 'fast';
  /**
   * Direction, valid only for effects that carry a `dir` attribute and only
   * within that effect's domain (validated on write):
   *   - `blinds`/`checker`/`comb`/`randomBar`: `horz` | `vert`
   *   - `push`/`wipe`: `l` | `r` | `u` | `d`
   *   - `cover`/`pull`: the above plus `lu` | `ru` | `ld` | `rd`
   *   - `strips`: `lu` | `ru` | `ld` | `rd`
   *   - `zoom`/`split`: `in` | `out`
   * A mismatched effect/direction pair throws; on any other effect a stray
   * `direction` is ignored.
   */
  direction?: string;
  /** For `wheel`: number of spokes (unsigned integer). Omitted means 4. */
  spokes?: number;
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
  /**
   * Effect duration in milliseconds (PowerPoint 2010's `p14:dur`). ECMA-376
   * only has the three `speed` steps, so the transition is written as
   * PowerPoint writes it: an `mc:AlternateContent` whose `p14` choice carries
   * the duration and whose fallback carries the nearest `speed`.
   */
  durationMs?: number;
}

/**
 * A transition read back from a deck. `effect` widens to `string` because a
 * file authored elsewhere can carry an effect element this library does not
 * model; writing one still requires a `TransitionEffect`.
 */
export type SlideTransition = Omit<TransitionOptions, 'effect'> & { effect: string };

// Per-effect `dir` value domains (ECMA-376 Part 1, pml.xsd). The effect
// element's CT type fixes which direction tokens are legal — they are NOT
// interchangeable: blinds wants horz/vert, push wants l/r/u/d, zoom wants
// in/out, etc. Emitting a token outside the effect's domain is schema-invalid,
// so we validate `direction` against the effect here (a boundary).
const DIR_ORIENT = new Set(['horz', 'vert']); // ST_Direction (CT_OrientationTransition)
const DIR_SIDE = new Set(['l', 'u', 'r', 'd']); // ST_TransitionSideDirectionType
const DIR_CORNER = new Set(['lu', 'ru', 'ld', 'rd']); // ST_TransitionCornerDirectionType
const DIR_EIGHT = new Set([...DIR_SIDE, ...DIR_CORNER]); // ST_TransitionEightDirectionType
const DIR_IN_OUT = new Set(['in', 'out']); // ST_TransitionInOutDirectionType
const DIR_DOMAINS: Readonly<Record<string, ReadonlySet<string>>> = {
  blinds: DIR_ORIENT,
  checker: DIR_ORIENT,
  comb: DIR_ORIENT,
  randomBar: DIR_ORIENT,
  push: DIR_SIDE,
  wipe: DIR_SIDE,
  cover: DIR_EIGHT,
  pull: DIR_EIGHT,
  strips: DIR_CORNER,
  zoom: DIR_IN_OUT,
  split: DIR_IN_OUT,
};
// Effects whose CT type carries `thruBlk` (CT_OptionalBlackTransition).
const THRU_BLK_EFFECTS = new Set(['fade', 'cut']);

// Returns the single effect child, or null for the "no transition effect"
// sentinel ('none' is not a valid effect element name — CT_SlideTransition's
// choice has no `none` member).
const buildEffectElement = (opts: TransitionOptions): XmlElement | null => {
  if (opts.effect === 'none') return null;
  const effect = oneOf(opts.effect, TRANSITION_EFFECTS, 'setSlideTransition: effect');
  const name = qname('p', effect, NS.pml);
  const attrs = [];
  if (opts.direction !== undefined) {
    // Only effects with a `dir` attribute carry a domain; for any other effect
    // a stray `direction` is ignored (it has nowhere valid to go).
    const domain = DIR_DOMAINS[opts.effect];
    if (domain !== undefined) {
      if (!domain.has(opts.direction)) {
        throw new Error(
          `setSlideTransition: direction "${opts.direction}" is not valid for effect ` +
            `"${opts.effect}" (allowed: ${[...domain].join(', ')})`,
        );
      }
      attrs.push(attr(ATTR_DIR, opts.direction));
    }
  }
  if (opts.spokes !== undefined && opts.effect === 'wheel') {
    if (!Number.isInteger(opts.spokes))
      throw new RangeError('setSlideTransition: spokes must be an integer');
    const spokes = boundedInt(opts.spokes, 'unsignedInt', 'setSlideTransition: spokes');
    attrs.push(attr(qname('', 'spokes', ''), String(spokes)));
  }
  // `orient` only exists on CT_SplitTransition.
  if (opts.orientation !== undefined && opts.effect === 'split') {
    oneOf(opts.orientation, ['horz', 'vert'], 'setSlideTransition: orientation');
    attrs.push(attr(ATTR_ORIENT, opts.orientation));
  }
  if (opts.thruBlack && THRU_BLK_EFFECTS.has(opts.effect)) {
    attrs.push(attr(ATTR_THRU_BLK, '1'));
  }
  return elem(name, { attrs });
};

/**
 * Returns the slide-level transition node: a `<p:transition>`, or the
 * `mc:AlternateContent` around two of them when a duration is set.
 * `sound` is the `<p:sndAc>` to keep (it follows the effect element).
 */
export const buildTransition = (
  opts: TransitionOptions,
  sound: XmlElement | null = null,
): XmlElement => {
  if (opts.speed !== undefined)
    oneOf(opts.speed, ['slow', 'med', 'fast'], 'setSlideTransition: speed');
  const duration =
    opts.durationMs === undefined
      ? undefined
      : unsignedIntMs(opts.durationMs, 'setSlideTransition: durationMs');
  const speed = opts.speed ?? (duration === undefined ? undefined : speedForDuration(duration));
  const attrs: XmlAttr[] = [];
  if (speed !== undefined) attrs.push(attr(ATTR_SPD, speed));
  if (opts.advanceOnClick === false) attrs.push(attr(ATTR_ADV_CLICK, '0'));
  if (opts.advanceAfterMs !== undefined) {
    // advTm is xsd:unsignedInt (0..4294967295 ms).
    const advTm = unsignedIntMs(opts.advanceAfterMs, 'setSlideTransition: advanceAfterMs');
    attrs.push(attr(ATTR_ADV_TM, String(advTm)));
  }
  const effect = buildEffectElement(opts);
  const transition = (extra: typeof attrs) =>
    elem(NAME_TRANSITION, {
      attrs: [...attrs, ...extra],
      children: [
        ...(effect === null ? [] : [cloneElement(effect)]),
        ...(sound === null ? [] : [cloneElement(sound)]),
      ],
    });
  if (duration === undefined) return transition([]);
  return elem(NAME_ALTERNATE_CONTENT, {
    prefixDecls: new Map([['mc', NS.mc]]),
    children: [
      elem(NAME_CHOICE, {
        attrs: [attr(ATTR_REQUIRES, 'p14')],
        prefixDecls: new Map([['p14', NS.p14]]),
        children: [transition([attr(ATTR_P14_DUR, String(duration))])],
      }),
      elem(NAME_FALLBACK, { children: [transition([])] }),
    ],
  });
};
