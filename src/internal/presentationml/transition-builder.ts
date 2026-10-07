// Builds the `<p:transition>` element that controls how PowerPoint
// animates from this slide to the next.
//
// Per ECMA-376 Part 1 §19.3.1.50 the transition carries:
//   - `spd` attribute: `slow` | `med` | `fast` (default `med`)
//   - `advClick` attribute: `1` to advance on click (default), `0` to disable
//   - `advTm` attribute: auto-advance time in milliseconds
//   - At most one child element from the effect catalog (fade, push,
//     cover, wipe, split, cut, dissolve, checker, blinds, randomBar,
//     zoom, circle, diamond, plus, wedge, ...).
//
// The effect attributes are not interchangeable: each effect element maps to a
// distinct CT type that only permits certain attributes (CT_OptionalBlackTransition
// → thruBlk; CT_SplitTransition → orient + dir; the direction families → dir).
// Emitting an attribute on an effect that doesn't allow it is schema-invalid, so
// buildEffectElement gates each attribute by the effect that accepts it.
//
// PowerPoint 2010 and later add effects outside ECMA-376 ([MS-PPTX] §2.2.1):
// nineteen in the p14 namespace (Vortex, Ripple, Shred, ...), the preset
// transitions of p15 (`p15:prstTrans`: Curtains, Origami, ...) and Morph in
// p159. They are written as PowerPoint writes them — an `mc:AlternateContent`
// whose choice requires the extension's namespace and whose fallback is the
// same transition with `<p:fade/>` ([MS-PPTX] §3.1), so a reader that knows
// only ECMA-376 still gets a valid fade.

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
const NAME_FADE = qname('p', 'fade', NS.pml);
const ATTR_SPD = qname('', 'spd', '');
const ATTR_ADV_CLICK = qname('', 'advClick', '');
const ATTR_ADV_TM = qname('', 'advTm', '');
const ATTR_DIR = qname('', 'dir', '');
const ATTR_ORIENT = qname('', 'orient', '');
const ATTR_THRU_BLK = qname('', 'thruBlk', '');
const ATTR_SPOKES = qname('', 'spokes', '');
const ATTR_PATTERN = qname('', 'pattern', '');
const ATTR_IS_CONTENT = qname('', 'isContent', '');
const ATTR_IS_INVERTED = qname('', 'isInverted', '');
const ATTR_HAS_BOUNCE = qname('', 'hasBounce', '');
const ATTR_PRST = qname('', 'prst', '');
const ATTR_INV_X = qname('', 'invX', '');
const ATTR_INV_Y = qname('', 'invY', '');
const ATTR_OPTION = qname('', 'option', '');
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
 * (ECMA-376 pml.xsd), emitted verbatim as `<p:{token}/>`.
 */
const ECMA_TRANSITION_EFFECTS = [
  'blinds',
  'checker',
  'circle',
  'dissolve',
  'comb',
  'cover',
  'cut',
  'diamond',
  'fade',
  'newsflash',
  'plus',
  'pull',
  'push',
  'random',
  'randomBar',
  'split',
  'strips',
  'wedge',
  'wheel',
  'wipe',
  'zoom',
] as const;

/**
 * The PowerPoint 2010 effects [MS-PPTX] §2.2.1 adds to that choice, in the
 * p14 namespace (`<p14:{token}/>`). Cube, Box, Rotate and Orbit are all
 * `prism`, told apart by `isContent` / `isInverted`; the gallery's Zoom is
 * `warp`; Clock's Counterclockwise is `wheelReverse`.
 */
const P14_TRANSITION_EFFECTS = [
  'conveyor',
  'doors',
  'ferris',
  'flash',
  'flip',
  'flythrough',
  'gallery',
  'glitter',
  'honeycomb',
  'pan',
  'prism',
  'reveal',
  'ripple',
  'shred',
  'switch',
  'vortex',
  'warp',
  'wheelReverse',
  'window',
] as const;

/**
 * Every effect element `setSlideTransition` writes, ECMA-376 and extensions
 * alike. The list is both the type's domain and the write-time validation
 * domain — keeping them one declaration stops them drifting apart.
 */
export const TRANSITION_EFFECTS = [
  ...ECMA_TRANSITION_EFFECTS,
  ...P14_TRANSITION_EFFECTS,
  // PowerPoint 2013's preset transitions (`p15:prstTrans`, [MS-PPTX] §2.4.3.8).
  'prstTrans',
  // PowerPoint 2016's Morph (`p159:morph`, [MS-PPTX] §2.6.1.1).
  'morph',
] as const;

/**
 * Transition effect token. Maps to the `<p:{token}/>` (or `<p14:…>`,
 * `<p15:prstTrans>`, `<p159:morph>`) child of `<p:transition>`, except
 * `'none'`: that is the library-level sentinel for "no effect element", which
 * the schema's choice has no member for.
 */
export type TransitionEffect = 'none' | (typeof TRANSITION_EFFECTS)[number];

/**
 * The `prst` names [MS-PPTX] §2.4.3.8 defines for `p15:prstTrans`. The
 * attribute is an `xsd:string` naming "the internal resource to use", so a
 * name outside this list is one PowerPoint has no transition for.
 */
export const TRANSITION_PRESETS = [
  'fallOver',
  'drape',
  'curtains',
  'wind',
  'prestige',
  'fracture',
  'crush',
  'peelOff',
  'pageCurlDouble',
  'pageCurlSingle',
  'airplane',
  'origami',
] as const;

/** A `p15:prstTrans` preset name. */
export type TransitionPreset = (typeof TRANSITION_PRESETS)[number];

/** ST_TransitionMorphOption ([MS-PPTX] §2.6.4.1): what Morph matches between slides. */
export type MorphOption = 'byObject' | 'byWord' | 'byChar';

export interface TransitionOptions {
  effect: TransitionEffect;
  /** Effect speed. Defaults to omitted (PowerPoint treats absence as `med`). */
  speed?: 'slow' | 'med' | 'fast';
  /**
   * Direction, valid only for effects that carry a `dir` attribute and only
   * within that effect's domain (validated on write). The side tokens name
   * the way the slide moves, so PowerPoint's "From Right" is `l`.
   *   - `blinds`/`checker`/`comb`/`randomBar`/`doors`/`window`: `horz` | `vert`
   *   - `push`/`wipe`/`vortex`/`pan`/`glitter`/`prism`: `l` | `r` | `u` | `d`
   *   - `cover`/`pull`: the above plus `lu` | `ru` | `ld` | `rd`
   *   - `strips`: `lu` | `ru` | `ld` | `rd`
   *   - `ripple`: `center` | `lu` | `ru` | `ld` | `rd`
   *   - `switch`/`flip`/`ferris`/`gallery`/`conveyor`/`reveal`: `l` | `r`
   *   - `zoom`/`split`/`warp`/`flythrough`/`shred`: `in` | `out`
   * A mismatched effect/direction pair throws; on any other effect a stray
   * `direction` is ignored.
   */
  direction?: string;
  /** For `wheel` and `wheelReverse`: number of spokes (unsigned integer). Omitted means 4. */
  spokes?: number;
  /** For `split`: orientation token (`horz` / `vert`). */
  orientation?: 'horz' | 'vert';
  /** For `fade`, `cut` and `reveal`: pass `true` to go through black. */
  thruBlack?: boolean;
  /**
   * For `glitter`: `diamond` (the default) or `hexagon`. For `shred`: `strip`
   * (the default) or `rectangle` — PowerPoint's "Strips" and "Particles".
   */
  pattern?: 'diamond' | 'hexagon' | 'strip' | 'rectangle';
  /**
   * For `prism`: the slide's content turns rather than the slide itself
   * (PowerPoint's Rotate and Orbit, against Cube and Box).
   */
  isContent?: boolean;
  /** For `prism`: the turn faces inwards (PowerPoint's Box and Orbit). */
  isInverted?: boolean;
  /** For `flythrough`: the slides bounce at the end. */
  hasBounce?: boolean;
  /** For `prstTrans`: which preset transition. Required for that effect. */
  preset?: TransitionPreset;
  /** For `prstTrans`: mirror the preset left to right ("Right" instead of "Left"). */
  invertX?: boolean;
  /** For `prstTrans`: mirror the preset top to bottom. */
  invertY?: boolean;
  /** For `morph`: what is matched between the slides. Defaults to `byObject`. */
  morphOption?: MorphOption;
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
 * model; writing one still requires a `TransitionEffect`. `preset` widens the
 * same way.
 */
export type SlideTransition = Omit<TransitionOptions, 'effect' | 'preset'> & {
  effect: string;
  preset?: string;
};

// Per-effect `dir` value domains (ECMA-376 Part 1 pml.xsd, [MS-PPTX] §5.1).
// The effect element's CT type fixes which direction tokens are legal — they
// are NOT interchangeable: blinds wants horz/vert, push wants l/r/u/d, zoom
// wants in/out, etc. Emitting a token outside the effect's domain is
// schema-invalid, so we validate `direction` against the effect here (a
// boundary).
const DIR_ORIENT = new Set(['horz', 'vert']); // ST_Direction (CT_OrientationTransition)
const DIR_SIDE = new Set(['l', 'u', 'r', 'd']); // ST_TransitionSideDirectionType
const DIR_CORNER = new Set(['lu', 'ru', 'ld', 'rd']); // ST_TransitionCornerDirectionType
const DIR_EIGHT = new Set([...DIR_SIDE, ...DIR_CORNER]); // ST_TransitionEightDirectionType
const DIR_IN_OUT = new Set(['in', 'out']); // ST_TransitionInOutDirectionType
const DIR_LEFT_RIGHT = new Set(['l', 'r']); // p14:ST_TransitionLeftRightDirectionType
// p14:ST_TransitionCornerAndCenterDirectionType
const DIR_CORNER_CENTER = new Set(['center', ...DIR_CORNER]);
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
  vortex: DIR_SIDE,
  pan: DIR_SIDE,
  glitter: DIR_SIDE,
  prism: DIR_SIDE,
  switch: DIR_LEFT_RIGHT,
  flip: DIR_LEFT_RIGHT,
  ferris: DIR_LEFT_RIGHT,
  gallery: DIR_LEFT_RIGHT,
  conveyor: DIR_LEFT_RIGHT,
  reveal: DIR_LEFT_RIGHT,
  ripple: DIR_CORNER_CENTER,
  doors: DIR_ORIENT,
  window: DIR_ORIENT,
  warp: DIR_IN_OUT,
  flythrough: DIR_IN_OUT,
  shred: DIR_IN_OUT,
};
// Effects whose CT type carries `thruBlk` (CT_OptionalBlackTransition, p14:CT_RevealTransition).
const THRU_BLK_EFFECTS = new Set(['fade', 'cut', 'reveal']);
// `pattern` domains: p14:ST_TransitionPattern and p14:ST_TransitionShredPattern.
const PATTERN_DOMAINS: Readonly<Record<string, readonly string[]>> = {
  glitter: ['diamond', 'hexagon'],
  shred: ['strip', 'rectangle'],
};
const MORPH_OPTIONS: readonly MorphOption[] = ['byObject', 'byWord', 'byChar'];

const P14_EFFECTS: ReadonlySet<string> = new Set(P14_TRANSITION_EFFECTS);

interface Extension {
  readonly ns: string;
  /** The prefix PowerPoint declares for it, which is also what `Requires` names. */
  readonly prefix: string;
}
const P14: Extension = { ns: NS.p14, prefix: 'p14' };
const P15: Extension = { ns: NS.p15, prefix: 'p15' };
const P159: Extension = { ns: NS.p159, prefix: 'p159' };

/** The extension namespace an effect element lives in; `null` for ECMA-376's own. */
const extensionOf = (effect: string): Extension | null =>
  P14_EFFECTS.has(effect) ? P14 : effect === 'prstTrans' ? P15 : effect === 'morph' ? P159 : null;

/**
 * The namespace URI of the element an effect token stands for, or `null` for
 * a token this library does not write. A reader uses it to tell `p14:flip`
 * from an element of the same local name in another namespace.
 */
export const transitionEffectNamespace = (effect: string): string | null =>
  (TRANSITION_EFFECTS as readonly string[]).includes(effect)
    ? (extensionOf(effect)?.ns ?? NS.pml)
    : null;

// xsd:boolean attributes whose schema default is false: written only when set.
const flag = (name: ReturnType<typeof qname>, on: boolean | undefined): XmlAttr[] =>
  on === true ? [attr(name, '1')] : [];

// Returns the single effect child, or null for the "no transition effect"
// sentinel ('none' is not a valid effect element name — CT_SlideTransition's
// choice has no `none` member).
const buildEffectElement = (opts: TransitionOptions): XmlElement | null => {
  if (opts.effect === 'none') return null;
  const effect = oneOf(opts.effect, TRANSITION_EFFECTS, 'setSlideTransition: effect');
  const extension = extensionOf(effect);
  const name = qname(extension?.prefix ?? 'p', effect, extension?.ns ?? NS.pml);
  const attrs: XmlAttr[] = [];
  if (opts.direction !== undefined) {
    // Only effects with a `dir` attribute carry a domain; for any other effect
    // a stray `direction` is ignored (it has nowhere valid to go).
    const domain = DIR_DOMAINS[effect];
    if (domain !== undefined) {
      if (!domain.has(opts.direction)) {
        throw new Error(
          `setSlideTransition: direction "${opts.direction}" is not valid for effect ` +
            `"${effect}" (allowed: ${[...domain].join(', ')})`,
        );
      }
      attrs.push(attr(ATTR_DIR, opts.direction));
    }
  }
  if (opts.spokes !== undefined && (effect === 'wheel' || effect === 'wheelReverse')) {
    if (!Number.isInteger(opts.spokes))
      throw new RangeError('setSlideTransition: spokes must be an integer');
    const spokes = boundedInt(opts.spokes, 'unsignedInt', 'setSlideTransition: spokes');
    attrs.push(attr(ATTR_SPOKES, String(spokes)));
  }
  // `orient` only exists on CT_SplitTransition.
  if (opts.orientation !== undefined && effect === 'split') {
    oneOf(opts.orientation, ['horz', 'vert'], 'setSlideTransition: orientation');
    attrs.push(attr(ATTR_ORIENT, opts.orientation));
  }
  if (THRU_BLK_EFFECTS.has(effect)) attrs.push(...flag(ATTR_THRU_BLK, opts.thruBlack));
  const patterns = PATTERN_DOMAINS[effect];
  if (opts.pattern !== undefined && patterns !== undefined) {
    attrs.push(attr(ATTR_PATTERN, oneOf(opts.pattern, patterns, 'setSlideTransition: pattern')));
  }
  if (effect === 'prism') {
    attrs.push(
      ...flag(ATTR_IS_CONTENT, opts.isContent),
      ...flag(ATTR_IS_INVERTED, opts.isInverted),
    );
  }
  if (effect === 'flythrough') attrs.push(...flag(ATTR_HAS_BOUNCE, opts.hasBounce));
  if (effect === 'prstTrans') {
    if (opts.preset === undefined)
      throw new Error('setSlideTransition: effect "prstTrans" needs a preset');
    const preset = oneOf(opts.preset, TRANSITION_PRESETS, 'setSlideTransition: preset');
    attrs.push(
      attr(ATTR_PRST, preset),
      ...flag(ATTR_INV_X, opts.invertX),
      ...flag(ATTR_INV_Y, opts.invertY),
    );
  }
  if (effect === 'morph') {
    // `option` is required on CT_MorphTransition; byObject is PowerPoint's "Objects".
    const option = oneOf(
      opts.morphOption ?? 'byObject',
      MORPH_OPTIONS,
      'setSlideTransition: morphOption',
    );
    attrs.push(attr(ATTR_OPTION, option));
  }
  return elem(name, { attrs });
};

/**
 * Returns the slide-level transition node: a `<p:transition>`, or the
 * `mc:AlternateContent` around two of them when a duration is set or the
 * effect is a PowerPoint extension. `sound` is the `<p:sndAc>` to keep (it
 * follows the effect element).
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
  const transition = (extra: readonly XmlAttr[], child: XmlElement | null) =>
    elem(NAME_TRANSITION, {
      attrs: [...attrs, ...extra],
      children: [
        ...(child === null ? [] : [cloneElement(child)]),
        ...(sound === null ? [] : [cloneElement(sound)]),
      ],
    });
  const extension = opts.effect === 'none' ? null : extensionOf(opts.effect);
  if (duration === undefined && extension === null) return transition([], effect);
  // The choice requires the namespace its effect lives in — p14 for an
  // ECMA-376 effect that only carries a duration. `p14:dur` is declared beside
  // a p15 or p159 effect too, because the duration is p14's attribute.
  const required = extension ?? P14;
  const prefixDecls = new Map([[required.prefix, required.ns]]);
  if (duration !== undefined) prefixDecls.set(P14.prefix, P14.ns);
  return elem(NAME_ALTERNATE_CONTENT, {
    prefixDecls: new Map([['mc', NS.mc]]),
    children: [
      elem(NAME_CHOICE, {
        attrs: [attr(ATTR_REQUIRES, required.prefix)],
        prefixDecls,
        children: [
          transition(duration === undefined ? [] : [attr(ATTR_P14_DUR, String(duration))], effect),
        ],
      }),
      // An extension effect falls back to a fade, as PowerPoint writes it.
      elem(NAME_FALLBACK, {
        children: [transition([], extension === null ? effect : elem(NAME_FADE))],
      }),
    ],
  });
};
