// Animation builder — emits a `<p:timing>` block carrying a single effect on
// one target shape. The caller (`setShapeAnimation`) merges it into the slide's
// existing tree, so a slide ends up with as many effects as there were calls.
//
// Scope:
//
//   - Entrance and exit presets — fade, fly, zoom and the transition filters
//     PowerPoint runs on one object (`<p:animEffect>`) — and one emphasis
//     preset (`spin`). Motion paths (`presetClass="path"`), the remaining
//     entrance / exit motions and the rest of the emphasis family are not
//     modelled.
//   - The whole shape, or one paragraph of its text body
//     (`<p:txEl><p:pRg>`, `<p:bldP build="p">`).
//
// The scaffolding around the behaviours — click stop, group wrapper, main
// sequence, `<p:bldLst>` — is fixed; the preset attributes, the node type, the
// delay, the target spid and the behaviours themselves are what vary.
//
// Preset numbers and behaviour shapes
// -----------------------------------
// ECMA-376 leaves `presetID` / `presetSubtype` as plain integers (§19.5.19,
// CT_TLCommonTimeNodeData) and `<p:attrName>` as a plain string: neither the
// preset catalogue nor the attribute vocabulary is in the standard. Both are
// PowerPoint conventions, so the tree each effect below writes is modelled on
// what PowerPoint 16 itself writes for the same gallery entry. The one place
// this file deliberately departs from PowerPoint is `fill="hold"` on an exit
// behaviour, which PowerPoint leaves off — see `buildAttrAnim`.
//
// What the attribute names mean is likewise convention: `ppt_x` / `ppt_y` are
// the shape's centre as a fraction of the slide, `ppt_w` / `ppt_h` its size as
// a fraction of the slide, and `r` its rotation. Cross-checked against
// LibreOffice's OOXML import, which maps them to X / Y / Width / Height /
// Rotate (`oox/source/ppt/pptfilterhelpers.cxx`, `getAttributeConversionList`).

import { oneOf, unsignedIntMs } from '../bounds.ts';
import { NS, type XmlElement, attr, elem, qname } from '../xml/index.ts';

const NAME_TIMING = qname('p', 'timing', NS.pml);
const NAME_TN_LST = qname('p', 'tnLst', NS.pml);
const NAME_PAR = qname('p', 'par', NS.pml);
const NAME_C_TN = qname('p', 'cTn', NS.pml);
const NAME_CHILD_TN_LST = qname('p', 'childTnLst', NS.pml);
const NAME_SEQ = qname('p', 'seq', NS.pml);
const NAME_ST_COND_LST = qname('p', 'stCondLst', NS.pml);
const NAME_PREV_COND_LST = qname('p', 'prevCondLst', NS.pml);
const NAME_NEXT_COND_LST = qname('p', 'nextCondLst', NS.pml);
const NAME_COND = qname('p', 'cond', NS.pml);
const NAME_TGT_EL = qname('p', 'tgtEl', NS.pml);
const NAME_SP_TGT = qname('p', 'spTgt', NS.pml);
const NAME_SLD_TGT = qname('p', 'sldTgt', NS.pml);
const NAME_ATTR_NAME_LST = qname('p', 'attrNameLst', NS.pml);
const NAME_ATTR_NAME_FN = qname('p', 'attrName', NS.pml);
const NAME_SET = qname('p', 'set', NS.pml);
const NAME_C_BHVR = qname('p', 'cBhvr', NS.pml);
const NAME_TO = qname('p', 'to', NS.pml);
const NAME_STR_VAL = qname('p', 'strVal', NS.pml);
const NAME_FLT_VAL = qname('p', 'fltVal', NS.pml);
const NAME_VAL = qname('p', 'val', NS.pml);
const NAME_ANIM = qname('p', 'anim', NS.pml);
const NAME_ANIM_ROT = qname('p', 'animRot', NS.pml);
const NAME_ANIM_EFFECT = qname('p', 'animEffect', NS.pml);
const NAME_TAV_LST = qname('p', 'tavLst', NS.pml);
const NAME_TAV = qname('p', 'tav', NS.pml);
const NAME_BLD_LST = qname('p', 'bldLst', NS.pml);
const NAME_BLD_P = qname('p', 'bldP', NS.pml);
const NAME_TX_EL = qname('p', 'txEl', NS.pml);
const NAME_P_RG = qname('p', 'pRg', NS.pml);

const ATTR_ID = qname('', 'id', '');
const ATTR_DUR = qname('', 'dur', '');
const ATTR_RESTART = qname('', 'restart', '');
const ATTR_NODE_TYPE = qname('', 'nodeType', '');
const ATTR_FILL = qname('', 'fill', '');
const ATTR_DELAY = qname('', 'delay', '');
const ATTR_PRESET_ID = qname('', 'presetID', '');
const ATTR_PRESET_CLASS = qname('', 'presetClass', '');
const ATTR_PRESET_SUBTYPE = qname('', 'presetSubtype', '');
const ATTR_GRP_ID = qname('', 'grpId', '');
const ATTR_CONCURRENT = qname('', 'concurrent', '');
const ATTR_NEXT_AC = qname('', 'nextAc', '');
const ATTR_SPID = qname('', 'spid', '');
const ATTR_EVT = qname('', 'evt', '');
const ATTR_VAL = qname('', 'val', '');
const ATTR_CALCMODE = qname('', 'calcmode', '');
const ATTR_VALUE_TYPE = qname('', 'valueType', '');
const ATTR_ADDITIVE = qname('', 'additive', '');
const ATTR_BY = qname('', 'by', '');
const ATTR_TM = qname('', 'tm', '');
const ATTR_BUILD = qname('', 'build', '');
const ATTR_ST = qname('', 'st', '');
const ATTR_END = qname('', 'end', '');
const ATTR_TRANSITION = qname('', 'transition', '');
const ATTR_FILTER = qname('', 'filter', '');

/**
 * What kind of effect to apply.
 *
 * Entrance effects put the shape on the slide and exit effects take it off.
 * `spin` is the one emphasis effect modelled here: it turns a shape that is
 * already on the slide and leaves it exactly where it was, so it never decides
 * whether the shape is shown.
 *
 * The `…In` / `…Out` pairs past the first four are PowerPoint's transition
 * filters applied to one object (`<p:animEffect filter>`, ECMA-376 §19.5.3):
 * Blinds, Checkerboard, Dissolve, Peek, Random Bars, Shape, Split, Strips,
 * Wedge, Wheel and Wipe.
 */
export type AnimationEffect =
  | 'appear'
  | 'fadeIn'
  | 'flyIn'
  | 'zoomIn'
  | 'blindsIn'
  | 'checkerboardIn'
  | 'dissolveIn'
  | 'peekIn'
  | 'randomBarsIn'
  | 'shapeIn'
  | 'splitIn'
  | 'stripsIn'
  | 'wedgeIn'
  | 'wheelIn'
  | 'wipeIn'
  | 'disappear'
  | 'fadeOut'
  | 'flyOut'
  | 'zoomOut'
  | 'blindsOut'
  | 'checkerboardOut'
  | 'dissolveOut'
  | 'peekOut'
  | 'randomBarsOut'
  | 'shapeOut'
  | 'splitOut'
  | 'stripsOut'
  | 'wedgeOut'
  | 'wheelOut'
  | 'wipeOut'
  | 'spin';

/**
 * Which edge of the slide a `flyIn` comes from, or a `flyOut` leaves by —
 * PowerPoint's "From Bottom" / Google Slides' "Fly in from bottom" — or, for
 * the four corners, the corner. A wipe or a peek names an edge too; strips
 * name a corner.
 */
export type AnimationDirection =
  | 'top'
  | 'right'
  | 'bottom'
  | 'left'
  | 'topLeft'
  | 'topRight'
  | 'bottomLeft'
  | 'bottomRight';

/** Which way blinds, random bars and a split run — and a checkerboard: across or down. */
export type AnimationOrientation = 'horizontal' | 'vertical';

/** Whether a shape or split effect opens from the centre outwards or closes inwards. */
export type AnimationInOut = 'in' | 'out';

/** The outline a `shapeIn` / `shapeOut` reveals the object through. */
export type AnimationShape = 'circle' | 'box' | 'diamond' | 'plus';

export const ANIMATION_EFFECTS: readonly AnimationEffect[] = [
  'appear',
  'fadeIn',
  'flyIn',
  'zoomIn',
  'blindsIn',
  'checkerboardIn',
  'dissolveIn',
  'peekIn',
  'randomBarsIn',
  'shapeIn',
  'splitIn',
  'stripsIn',
  'wedgeIn',
  'wheelIn',
  'wipeIn',
  'disappear',
  'fadeOut',
  'flyOut',
  'zoomOut',
  'blindsOut',
  'checkerboardOut',
  'dissolveOut',
  'peekOut',
  'randomBarsOut',
  'shapeOut',
  'splitOut',
  'stripsOut',
  'wedgeOut',
  'wheelOut',
  'wipeOut',
  'spin',
];

const SIDES: readonly AnimationDirection[] = ['top', 'right', 'bottom', 'left'];
const CORNERS: readonly AnimationDirection[] = ['topLeft', 'topRight', 'bottomLeft', 'bottomRight'];

export const ANIMATION_DIRECTIONS: readonly AnimationDirection[] = [...SIDES, ...CORNERS];

const ORIENTATIONS: readonly AnimationOrientation[] = ['horizontal', 'vertical'];
const IN_OUT: readonly AnimationInOut[] = ['in', 'out'];
const SHAPES: readonly AnimationShape[] = ['circle', 'box', 'diamond', 'plus'];
/** The spoke counts PowerPoint's Wheel offers (1, 2, 3, 4 and 8 Spokes). */
const WHEEL_SPOKES: readonly number[] = [1, 2, 3, 4, 8];

/**
 * The `presetSubtype` PowerPoint writes for a directional preset. It is a
 * bitmask over the four edges, so a corner is the two bits together
 * (top-left is 1|8 = 9).
 *
 * The bit names the *edge*, not the way the shape travels: subtype 4 is the
 * bottom edge for both classes — an entrance rises from below it, an exit
 * drops through it.
 */
const DIRECTION_SUBTYPES: Record<AnimationDirection, number> = {
  top: 1,
  right: 2,
  bottom: 4,
  left: 8,
  topLeft: 9,
  topRight: 3,
  bottomLeft: 12,
  bottomRight: 6,
};
// Orientation and in / out are bits of the same mask: 10 is left|right, 5 is
// top|bottom, 16 in and 32 out.
const ORIENTATION_SUBTYPES: Record<AnimationOrientation, number> = { horizontal: 10, vertical: 5 };
const IN_OUT_SUBTYPES: Record<AnimationInOut, number> = { in: 16, out: 32 };

/** How an effect animates, beyond putting the shape on the slide or off it. */
type Motion = 'none' | 'fade' | 'fly' | 'zoom' | 'spin' | 'filter' | 'peek';

/** The options an effect's family takes, besides timing. */
export type AnimationOptionName = 'direction' | 'orientation' | 'inOut' | 'shape' | 'spokes';

/** Every option an effect can carry, resolved: `null` where the effect takes none. */
export interface AnimationEffectOptions {
  readonly direction: AnimationDirection | null;
  readonly orientation: AnimationOrientation | null;
  readonly inOut: AnimationInOut | null;
  readonly shape: AnimationShape | null;
  readonly spokes: number | null;
}

/** The values each option an effect takes may have, its default first. */
export interface AnimationOptionDomains {
  readonly direction?: readonly AnimationDirection[];
  readonly orientation?: readonly AnimationOrientation[];
  readonly inOut?: readonly AnimationInOut[];
  readonly shape?: readonly AnimationShape[];
  readonly spokes?: readonly number[];
}

interface Family {
  readonly motion: Motion;
  /** The values each option this family takes may have; the first is the default. */
  readonly domains: AnimationOptionDomains;
  /** PowerPoint's own default length for the preset, in milliseconds. */
  readonly durationMs: number;
  readonly presetId: (o: AnimationEffectOptions) => number;
  readonly presetSubtype: (o: AnimationEffectOptions) => number;
  /** The `<p:animEffect filter>` an entrance (`entering`) or exit writes. */
  readonly filter?: (o: AnimationEffectOptions, entering: boolean) => string;
}

// What PowerPoint writes for each preset — preset id, subtype and filter —
// as its object model records them (MsoAnimEffect, PowerPoint 16; see
// POWERPOINT_PARITY.md "Animations"). Durations are PowerPoint's defaults for
// the preset, except that the four original effects and Spin keep this
// library's 500 ms.
const constant = (value: number) => (): number => value;
const edgeWord: Record<string, string> = {
  top: 'up',
  right: 'right',
  bottom: 'down',
  left: 'left',
};
const oppositeEdge: Record<string, string> = {
  top: 'bottom',
  right: 'left',
  bottom: 'top',
  left: 'right',
};
const cornerWord: Record<string, string> = {
  topLeft: 'upLeft',
  topRight: 'upRight',
  bottomLeft: 'downLeft',
  bottomRight: 'downRight',
};
const SHAPE_PRESET_IDS: Record<AnimationShape, number> = {
  circle: 6,
  box: 4,
  diamond: 8,
  plus: 13,
};
const capitalised = (word: string): string => word[0]!.toUpperCase() + word.slice(1);

const FAMILIES = {
  appear: {
    motion: 'none',
    domains: {},
    durationMs: 500,
    presetId: constant(1),
    presetSubtype: constant(0),
  },
  fade: {
    motion: 'fade',
    domains: {},
    durationMs: 500,
    presetId: constant(10),
    presetSubtype: constant(0),
  },
  fly: {
    motion: 'fly',
    domains: { direction: ['bottom', ...SIDES.filter((d) => d !== 'bottom'), ...CORNERS] },
    durationMs: 500,
    presetId: constant(2),
    presetSubtype: (o) => DIRECTION_SUBTYPES[o.direction!],
  },
  zoom: {
    motion: 'zoom',
    domains: {},
    durationMs: 500,
    presetId: constant(23),
    presetSubtype: constant(0),
  },
  spin: {
    motion: 'spin',
    domains: {},
    durationMs: 500,
    presetId: constant(8),
    presetSubtype: constant(0),
  },
  blinds: {
    motion: 'filter',
    domains: { orientation: ORIENTATIONS },
    durationMs: 500,
    presetId: constant(3),
    presetSubtype: (o) => ORIENTATION_SUBTYPES[o.orientation!],
    filter: (o) => `blinds(${o.orientation})`,
  },
  checkerboard: {
    motion: 'filter',
    domains: { orientation: ORIENTATIONS },
    durationMs: 500,
    presetId: constant(5),
    presetSubtype: (o) => ORIENTATION_SUBTYPES[o.orientation!],
    filter: (o) => `checkerboard(${o.orientation === 'horizontal' ? 'across' : 'down'})`,
  },
  dissolve: {
    motion: 'filter',
    domains: {},
    durationMs: 500,
    presetId: constant(9),
    presetSubtype: constant(0),
    filter: () => 'dissolve',
  },
  peek: {
    motion: 'peek',
    domains: { direction: ['bottom', ...SIDES.filter((d) => d !== 'bottom')] },
    durationMs: 500,
    presetId: constant(12),
    presetSubtype: (o) => DIRECTION_SUBTYPES[o.direction!],
    // The object slides in from the edge while the wipe uncovers it from the
    // far side, so an entrance wipes towards the opposite edge.
    filter: (o, entering) =>
      `wipe(${edgeWord[entering ? oppositeEdge[o.direction!]! : o.direction!]})`,
  },
  randomBars: {
    motion: 'filter',
    domains: { orientation: ORIENTATIONS },
    durationMs: 500,
    presetId: constant(14),
    presetSubtype: (o) => ORIENTATION_SUBTYPES[o.orientation!],
    // ECMA-376's filter table spells this `randomBars(…)`; PowerPoint writes
    // and reads `randombar(…)`, so that is what is written.
    filter: (o) => `randombar(${o.orientation})`,
  },
  shape: {
    motion: 'filter',
    domains: { shape: SHAPES, inOut: IN_OUT },
    durationMs: 2000,
    presetId: (o) => SHAPE_PRESET_IDS[o.shape!],
    presetSubtype: (o) => IN_OUT_SUBTYPES[o.inOut!],
    filter: (o) => `${o.shape}(${o.inOut})`,
  },
  split: {
    motion: 'filter',
    domains: { orientation: ['vertical', 'horizontal'], inOut: IN_OUT },
    durationMs: 500,
    presetId: constant(16),
    presetSubtype: (o) =>
      IN_OUT_SUBTYPES[o.inOut!] +
      (o.orientation === 'vertical' ? 5 : ORIENTATION_SUBTYPES.horizontal),
    filter: (o) => `barn(${o.inOut}${capitalised(o.orientation!)})`,
  },
  strips: {
    motion: 'filter',
    domains: { direction: ['bottomLeft', 'topLeft', 'topRight', 'bottomRight'] },
    durationMs: 500,
    presetId: constant(18),
    presetSubtype: (o) => DIRECTION_SUBTYPES[o.direction!],
    filter: (o) => `strips(${cornerWord[o.direction!]})`,
  },
  wedge: {
    motion: 'filter',
    domains: {},
    durationMs: 2000,
    presetId: constant(20),
    presetSubtype: constant(0),
    filter: () => 'wedge',
  },
  wheel: {
    motion: 'filter',
    domains: { spokes: WHEEL_SPOKES },
    durationMs: 2000,
    presetId: constant(21),
    presetSubtype: (o) => o.spokes!,
    filter: (o) => `wheel(${o.spokes})`,
  },
  wipe: {
    motion: 'filter',
    domains: { direction: ['bottom', ...SIDES.filter((d) => d !== 'bottom')] },
    durationMs: 500,
    presetId: constant(22),
    presetSubtype: (o) => DIRECTION_SUBTYPES[o.direction!],
    filter: (o) => `wipe(${edgeWord[o.direction!]})`,
  },
} satisfies Record<string, Family>;

type FamilyName = keyof typeof FAMILIES;

interface PresetDescriptor {
  readonly presetClass: 'entr' | 'exit' | 'emph';
  readonly family: FamilyName;
  /** A subtype the preset writes whatever its options (zoom's 16 and 32). */
  readonly presetSubtype?: number;
}

const PRESETS: Record<AnimationEffect, PresetDescriptor> = {
  appear: { presetClass: 'entr', family: 'appear' },
  fadeIn: { presetClass: 'entr', family: 'fade' },
  flyIn: { presetClass: 'entr', family: 'fly' },
  zoomIn: { presetClass: 'entr', family: 'zoom', presetSubtype: 16 },
  blindsIn: { presetClass: 'entr', family: 'blinds' },
  checkerboardIn: { presetClass: 'entr', family: 'checkerboard' },
  dissolveIn: { presetClass: 'entr', family: 'dissolve' },
  peekIn: { presetClass: 'entr', family: 'peek' },
  randomBarsIn: { presetClass: 'entr', family: 'randomBars' },
  shapeIn: { presetClass: 'entr', family: 'shape' },
  splitIn: { presetClass: 'entr', family: 'split' },
  stripsIn: { presetClass: 'entr', family: 'strips' },
  wedgeIn: { presetClass: 'entr', family: 'wedge' },
  wheelIn: { presetClass: 'entr', family: 'wheel' },
  wipeIn: { presetClass: 'entr', family: 'wipe' },
  disappear: { presetClass: 'exit', family: 'appear' },
  fadeOut: { presetClass: 'exit', family: 'fade' },
  flyOut: { presetClass: 'exit', family: 'fly' },
  zoomOut: { presetClass: 'exit', family: 'zoom', presetSubtype: 32 },
  blindsOut: { presetClass: 'exit', family: 'blinds' },
  checkerboardOut: { presetClass: 'exit', family: 'checkerboard' },
  dissolveOut: { presetClass: 'exit', family: 'dissolve' },
  peekOut: { presetClass: 'exit', family: 'peek' },
  randomBarsOut: { presetClass: 'exit', family: 'randomBars' },
  shapeOut: { presetClass: 'exit', family: 'shape' },
  splitOut: { presetClass: 'exit', family: 'split' },
  stripsOut: { presetClass: 'exit', family: 'strips' },
  wedgeOut: { presetClass: 'exit', family: 'wedge' },
  wheelOut: { presetClass: 'exit', family: 'wheel' },
  wipeOut: { presetClass: 'exit', family: 'wipe' },
  spin: { presetClass: 'emph', family: 'spin' },
};

const familyOf = (effect: AnimationEffect): Family => FAMILIES[PRESETS[effect].family];

/** The options `effect` takes and the values each may have, the default first. */
export const animationOptionDomains = (effect: AnimationEffect): AnimationOptionDomains =>
  familyOf(effect).domains;

// The behaviour elements each motion writes. Appear, fade, fly and zoom share
// one list: a node naming one of them and holding a `<p:set>` and `<p:anim>`s
// is rewritable whichever it is.
const BEHAVIOURS_BY_MOTION: Record<Motion, ReadonlySet<string>> = {
  none: new Set(['set', 'anim']),
  fade: new Set(['set', 'anim']),
  fly: new Set(['set', 'anim']),
  zoom: new Set(['set', 'anim']),
  spin: new Set(['animRot']),
  filter: new Set(['set', 'animEffect']),
  peek: new Set(['set', 'anim', 'animEffect']),
};

/** The local names of the behaviour elements this library writes for `effect`. */
export const effectBehaviourNames = (effect: AnimationEffect): ReadonlySet<string> =>
  BEHAVIOURS_BY_MOTION[familyOf(effect).motion];

/** Whether the effect flies, and so takes a direction. */
export const isDirectionalEffect = (effect: AnimationEffect): boolean =>
  PRESETS[effect].family === 'fly';

/** PowerPoint's own default length for the preset, in milliseconds. */
export const defaultAnimationDurationMs = (effect: AnimationEffect): number =>
  familyOf(effect).durationMs;

/** The options passed for an effect, before they are checked against it. */
export type AnimationOptionValues = Partial<Record<AnimationOptionName, unknown>>;

/**
 * Validates the options passed for `effect` and fills in the defaults of the
 * ones it takes. An option the effect does not take is an error rather than a
 * no-op: it would otherwise read as something the file never records.
 */
export const resolveAnimationOptions = (
  effect: AnimationEffect,
  given: AnimationOptionValues,
  label: string,
): AnimationEffectOptions => {
  const domains = familyOf(effect).domains;
  const pick = <T extends string | number>(
    name: AnimationOptionName,
    domain: readonly T[] | undefined,
  ): T | null => {
    const value = given[name];
    if (domain === undefined) {
      if (value === undefined) return null;
      throw new RangeError(
        `${label}: ${name} only applies to ${ANIMATION_EFFECTS.filter(
          (other) => familyOf(other).domains[name] !== undefined,
        ).join(', ')}; ${JSON.stringify(effect)} does not take it.`,
      );
    }
    if (value === undefined) return domain[0]!;
    const found = domain.find((candidate) => candidate === value);
    if (found === undefined) {
      throw new RangeError(
        `${label}: ${name} ${JSON.stringify(value)} is not one of ${domain.join(', ')} for ${JSON.stringify(effect)}.`,
      );
    }
    return found;
  };
  return {
    direction: pick('direction', domains.direction),
    orientation: pick('orientation', domains.orientation),
    inOut: pick('inOut', domains.inOut),
    shape: pick('shape', domains.shape),
    spokes: pick('spokes', domains.spokes),
  };
};

/** One `(presetClass, presetID, presetSubtype)` triple this library writes, and what it means. */
export interface AnimationPresetEntry {
  readonly presetClass: 'entr' | 'exit' | 'emph';
  readonly presetId: number;
  readonly presetSubtype: number;
  readonly effect: AnimationEffect;
  readonly options: AnimationEffectOptions;
  /**
   * Whether the effect has no options at all, so a tree that leaves the
   * subtype out can only have meant this one.
   */
  readonly optionless: boolean;
}

// Every combination of the options a family takes; one empty set when it takes none.
const combinations = (d: AnimationOptionDomains): AnimationEffectOptions[] => {
  const axis = <T>(domain: readonly T[] | undefined): readonly (T | null)[] => domain ?? [null];
  return axis(d.direction).flatMap((direction) =>
    axis(d.orientation).flatMap((orientation) =>
      axis(d.inOut).flatMap((inOut) =>
        axis(d.shape).flatMap((shape) =>
          axis(d.spokes).map((spokes) => ({ direction, orientation, inOut, shape, spokes })),
        ),
      ),
    ),
  );
};

/**
 * Every preset triple this library writes, with the effect and options it
 * stands for — the one table both the writer and the reader go by.
 */
export const ANIMATION_PRESET_ENTRIES: readonly AnimationPresetEntry[] = ANIMATION_EFFECTS.flatMap(
  (effect) => {
    const preset = PRESETS[effect];
    const family = familyOf(effect);
    return combinations(family.domains).map((options) => ({
      presetClass: preset.presetClass,
      presetId: family.presetId(options),
      presetSubtype: preset.presetSubtype ?? family.presetSubtype(options),
      effect,
      options,
      optionless: Object.keys(family.domains).length === 0 && preset.presetSubtype === undefined,
    }));
  },
);

/** A full turn in `a:ST_Angle`, which counts sixtieth-thousandths of a degree. */
export const FULL_TURN = 360 * 60000;

/**
 * Where the `<p:set>` that ends an exit stands, measured from the start of the
 * effect. The shape has to stay on the slide until its motion is over, and the
 * kick itself takes the one millisecond PowerPoint gives it.
 */
export const trailingHideDelayMs = (durationMs: number): number => Math.max(0, durationMs - 1);

/** When an effect runs relative to the one before it. */
export type AnimationStartCondition = 'click' | 'withPrevious' | 'afterPrevious';

// ST_TLTimeNodeType tokens for the three start conditions.
const START_NODE_TYPES: Record<AnimationStartCondition, string> = {
  click: 'clickEffect',
  withPrevious: 'withEffect',
  afterPrevious: 'afterEffect',
};

export interface AnimationOptions {
  /** Which preset effect to apply. */
  readonly effect: AnimationEffect;
  /**
   * Which edge of the slide a `'flyIn'` comes from, or a `'flyOut'` leaves by
   * — any of the eight. `'wipeIn'`, `'wipeOut'`, `'peekIn'` and `'peekOut'`
   * take one of the four edges, `'stripsIn'` and `'stripsOut'` one of the four
   * corners. Defaults to PowerPoint's own default for the preset (`'bottom'`;
   * `'bottomLeft'` for strips). Passing it for an effect that takes no
   * direction is an error rather than a no-op: it would otherwise read as a
   * direction the file never records. The same holds for the options below.
   */
  readonly direction?: AnimationDirection;
  /**
   * For blinds, random bars and split: which way the bars run. For a
   * checkerboard, `'horizontal'` is PowerPoint's "Across" and `'vertical'` its
   * "Down". Defaults to `'horizontal'` (`'vertical'` for split).
   */
  readonly orientation?: AnimationOrientation;
  /** For shape and split: open from the centre (`'out'`) or close in on it (`'in'`). Defaults to `'in'`. */
  readonly inOut?: AnimationInOut;
  /** For `'shapeIn'` / `'shapeOut'`: the outline. Defaults to `'circle'`. */
  readonly shape?: AnimationShape;
  /** For `'wheelIn'` / `'wheelOut'`: 1, 2, 3, 4 or 8 spokes. Defaults to 1. */
  readonly spokes?: number;
  /**
   * Animation length in milliseconds. Defaults to PowerPoint's default for
   * the preset — 2000ms for shape, wedge and wheel, 500ms for the rest.
   * `'appear'` and `'disappear'` are instantaneous by definition of the preset
   * and write no timed behaviour at all, so it does not reach them.
   */
  readonly durationMs?: number;
  /**
   * What starts the effect. Defaults to `'click'` — the effect waits for the
   * viewer's next click. `'withPrevious'` runs it alongside the effect before
   * it, `'afterPrevious'` once that effect has finished. When the effect is
   * the slide's first, neither has a predecessor to follow, so both start as
   * soon as the slide appears instead of waiting for a click.
   */
  readonly start?: AnimationStartCondition;
  /** How long to wait once the start condition is met. Defaults to 0ms. */
  readonly delayMs?: number;
  /**
   * Reveal the shape's text one paragraph at a time instead of animating the
   * shape as a whole — PowerPoint's "By paragraph", Google Slides' "By
   * paragraph". Each paragraph gets its own effect with the same `start`, so
   * the default `'click'` advances one paragraph per click.
   */
  readonly byParagraph?: boolean;
}

/**
 * The `<p:tgtEl>` an effect's behaviours animate. A paragraph index narrows it
 * to that one paragraph (`<p:txEl><p:pRg>`, CT_IndexRange — inclusive on both
 * ends, so a single paragraph has `st` and `end` equal).
 */
const buildTarget = (spid: number, paragraph: number | null): XmlElement => {
  const spTgt = elem(NAME_SP_TGT, {
    attrs: [attr(ATTR_SPID, String(spid))],
    children:
      paragraph === null
        ? []
        : [
            elem(NAME_TX_EL, {
              children: [
                elem(NAME_P_RG, {
                  attrs: [attr(ATTR_ST, String(paragraph)), attr(ATTR_END, String(paragraph))],
                }),
              ],
            }),
          ],
  });
  return elem(NAME_TGT_EL, { children: [spTgt] });
};

const buildSetVisibility = (
  spid: number,
  id: number,
  visible: boolean,
  delayMs: number,
  paragraph: number | null,
): XmlElement => {
  const tgt = buildTarget(spid, paragraph);
  const attrName = elem(NAME_ATTR_NAME_LST, {
    children: [elem(NAME_ATTR_NAME_FN, { children: [{ kind: 'text', data: 'style.visibility' }] })],
  });
  const cTn = elem(NAME_C_TN, {
    attrs: [attr(ATTR_ID, String(id)), attr(ATTR_DUR, '1'), attr(ATTR_FILL, 'hold')],
    children: [
      elem(NAME_ST_COND_LST, {
        children: [elem(NAME_COND, { attrs: [attr(ATTR_DELAY, String(delayMs))] })],
      }),
    ],
  });
  const cBhvr = elem(NAME_C_BHVR, { children: [cTn, tgt, attrName] });
  const to = elem(NAME_TO, {
    children: [elem(NAME_STR_VAL, { attrs: [attr(ATTR_VAL, visible ? 'visible' : 'hidden')] })],
  });
  return elem(NAME_SET, { children: [cBhvr, to] });
};

/**
 * One end of an animated value. A plain number is a `<p:fltVal>`; anything
 * naming the shape's own geometry is a `<p:strVal>` holding PowerPoint's
 * formula language — `#ppt_h` and `ppt_h` are the shape's height as a fraction
 * of the slide, so `1+#ppt_h/2` is the centre it needs to sit at to be just
 * below the slide's bottom edge.
 */
type AnimValue = { readonly flt: string } | { readonly str: string };

const valueElement = (value: AnimValue): XmlElement =>
  elem(NAME_VAL, {
    children: [
      'flt' in value
        ? elem(NAME_FLT_VAL, { attrs: [attr(ATTR_VAL, value.flt)] })
        : elem(NAME_STR_VAL, { attrs: [attr(ATTR_VAL, value.str)] }),
    ],
  });

interface AttrAnimSpec {
  readonly spid: number;
  readonly id: number;
  readonly durationMs: number;
  readonly attrName: string;
  readonly from: AnimValue;
  readonly to: AnimValue;
  /**
   * `additive="base"` composes the animated value onto the shape's authored
   * one. PowerPoint writes it for the effects whose keyframes refer back to
   * that value (`#ppt_x`) and leaves it off the ones that state absolute
   * fractions, so this mirrors the preset it came from.
   */
  readonly additive: boolean;
  readonly paragraph: number | null;
}

const buildAttrAnim = (spec: AttrAnimSpec): XmlElement => {
  const attrName = elem(NAME_ATTR_NAME_LST, {
    children: [elem(NAME_ATTR_NAME_FN, { children: [{ kind: 'text', data: spec.attrName }] })],
  });
  // PowerPoint leaves `fill` off an exit behaviour, because the `<p:set>` that
  // hides the shape a millisecond before the end settles what is seen either
  // way. We state `hold` on every behaviour instead: the attribute has no
  // default in the schema, so an omitted one says nothing about what becomes of
  // the value, and a reader then cannot tell a held value from an unknown one.
  const cTn = elem(NAME_C_TN, {
    attrs: [
      attr(ATTR_ID, String(spec.id)),
      attr(ATTR_DUR, String(spec.durationMs)),
      attr(ATTR_FILL, 'hold'),
    ],
  });
  const cBhvr = elem(NAME_C_BHVR, {
    attrs: spec.additive ? [attr(ATTR_ADDITIVE, 'base')] : [],
    children: [cTn, buildTarget(spec.spid, spec.paragraph), attrName],
  });
  const tavLst = elem(NAME_TAV_LST, {
    children: [
      elem(NAME_TAV, { attrs: [attr(ATTR_TM, '0')], children: [valueElement(spec.from)] }),
      elem(NAME_TAV, { attrs: [attr(ATTR_TM, '100000')], children: [valueElement(spec.to)] }),
    ],
  });
  return elem(NAME_ANIM, {
    attrs: [attr(ATTR_CALCMODE, 'lin'), attr(ATTR_VALUE_TYPE, 'num')],
    children: [cBhvr, tavLst],
  });
};

/**
 * The two position behaviours of a fly. Both axes are written even though only
 * one of them moves: PowerPoint holds the other one steady rather than leaving
 * it to whatever else is animating the shape, and an effect that states both is
 * the one its preset id stands for.
 *
 * `#` marks the shape's own authored value. Entrance keyframes carry it and
 * exit keyframes do not, which is the difference between the two preset rows.
 */
const buildFlyAnims = (
  spid: number,
  firstId: number,
  durationMs: number,
  entering: boolean,
  direction: AnimationDirection,
  paragraph: number | null,
): XmlElement[] => {
  const base = (axis: 'x' | 'y' | 'w' | 'h'): string => `${entering ? '#' : ''}ppt_${axis}`;
  // The edge the direction names on each axis; a corner names one on both.
  const edge = { x: horizontalEdge(direction), y: verticalEdge(direction) };
  // Just outside the edge: the shape's centre one half-size beyond it.
  const offSlide = (axis: 'x' | 'y'): string => {
    const half = `${base(axis === 'x' ? 'w' : 'h')}/2`;
    const far = edge[axis] === 'right' || edge[axis] === 'bottom';
    return far ? `1+${half}` : `0-${half}`;
  };
  const moves = (axis: 'x' | 'y'): boolean => edge[axis] !== null;
  return (['x', 'y'] as const).map((axis, at) => {
    const own: AnimValue = { str: base(axis) };
    const away: AnimValue = { str: offSlide(axis) };
    const still = !moves(axis);
    return buildAttrAnim({
      spid,
      id: firstId + at,
      durationMs,
      attrName: `ppt_${axis}`,
      from: still || !entering ? own : away,
      to: still || entering ? own : away,
      additive: true,
      paragraph,
    });
  });
};

const horizontalEdge = (direction: AnimationDirection): 'left' | 'right' | null =>
  direction === 'left' || direction === 'topLeft' || direction === 'bottomLeft'
    ? 'left'
    : direction === 'right' || direction === 'topRight' || direction === 'bottomRight'
      ? 'right'
      : null;

const verticalEdge = (direction: AnimationDirection): 'top' | 'bottom' | null =>
  direction === 'top' || direction === 'topLeft' || direction === 'topRight'
    ? 'top'
    : direction === 'bottom' || direction === 'bottomLeft' || direction === 'bottomRight'
      ? 'bottom'
      : null;

/**
 * The slide of a peek: the shape travels 1.125 of its own size from just past
 * the edge into place (or out again), on the one axis the edge is on. Unlike a
 * fly, PowerPoint writes only the moving axis, and keeps the `#` on both ends
 * of an exit too.
 */
const buildPeekAnim = (
  spid: number,
  id: number,
  durationMs: number,
  entering: boolean,
  direction: AnimationDirection,
  paragraph: number | null,
): XmlElement => {
  const horizontal = horizontalEdge(direction) !== null;
  const axis = horizontal ? 'x' : 'y';
  const size = horizontal ? 'w' : 'h';
  const sign = direction === 'right' || direction === 'bottom' ? '+' : '-';
  const own: AnimValue = { str: `#ppt_${axis}` };
  const away: AnimValue = { str: `#ppt_${axis}${sign}#ppt_${size}*1.125000` };
  return buildAttrAnim({
    spid,
    id,
    durationMs,
    attrName: `ppt_${axis}`,
    from: entering ? away : own,
    to: entering ? own : away,
    additive: true,
    paragraph,
  });
};

/**
 * A transition filter applied to the one object (ECMA-376 §19.5.3,
 * `<p:animEffect>`): `in` uncovers the shape through it, `out` covers it up.
 */
const buildFilterEffect = (
  spid: number,
  id: number,
  durationMs: number,
  entering: boolean,
  filter: string,
  paragraph: number | null,
): XmlElement => {
  // `fill="hold"` for the same reason as `buildAttrAnim`: PowerPoint leaves it
  // off, but an omitted fill says nothing about the end state.
  const cTn = elem(NAME_C_TN, {
    attrs: [attr(ATTR_ID, String(id)), attr(ATTR_DUR, String(durationMs)), attr(ATTR_FILL, 'hold')],
  });
  return elem(NAME_ANIM_EFFECT, {
    attrs: [attr(ATTR_TRANSITION, entering ? 'in' : 'out'), attr(ATTR_FILTER, filter)],
    children: [elem(NAME_C_BHVR, { children: [cTn, buildTarget(spid, paragraph)] })],
  });
};

/**
 * The two size behaviours of a zoom: the shape grows from nothing, or shrinks
 * to it. PowerPoint drives `ppt_w` / `ppt_h` rather than `<p:animScale>` for
 * this preset, so the shape scales about the centre `ppt_x` / `ppt_y` names.
 */
const buildZoomAnims = (
  spid: number,
  firstId: number,
  durationMs: number,
  entering: boolean,
  paragraph: number | null,
): XmlElement[] =>
  (['w', 'h'] as const).map((axis, at) => {
    const own: AnimValue = { str: `${entering ? '#' : ''}ppt_${axis}` };
    const none: AnimValue = { flt: '0' };
    return buildAttrAnim({
      spid,
      id: firstId + at,
      durationMs,
      attrName: `ppt_${axis}`,
      from: entering ? none : own,
      to: entering ? own : none,
      additive: false,
      paragraph,
    });
  });

/** A full clockwise turn about the shape's centre, which is what `r` names. */
const buildSpinAnim = (
  spid: number,
  id: number,
  durationMs: number,
  paragraph: number | null,
): XmlElement => {
  const attrName = elem(NAME_ATTR_NAME_LST, {
    children: [elem(NAME_ATTR_NAME_FN, { children: [{ kind: 'text', data: 'r' }] })],
  });
  const cTn = elem(NAME_C_TN, {
    attrs: [attr(ATTR_ID, String(id)), attr(ATTR_DUR, String(durationMs)), attr(ATTR_FILL, 'hold')],
  });
  const cBhvr = elem(NAME_C_BHVR, {
    children: [cTn, buildTarget(spid, paragraph), attrName],
  });
  return elem(NAME_ANIM_ROT, {
    attrs: [attr(ATTR_BY, String(FULL_TURN))],
    children: [cBhvr],
  });
};

/**
 * Builds the complete `<p:timing>` element for a single effect on the given
 * shape id. The result is a standalone tree whose outermost `<p:par>` is one
 * click stop; merging it behind effects that already exist is the caller's
 * job, and only the caller knows whether that stop survives.
 *
 * `paragraph` narrows the effect to one paragraph of the shape's text body.
 * A by-paragraph build is one such effect per paragraph, which is why the
 * caller loops rather than this function: the effects share a build group and
 * each needs merging into whatever the slide holds by then.
 */
export const buildSingleEffectTiming = (
  spid: number,
  opts: AnimationOptions,
  ctx: { readonly paragraph?: number | null; readonly label?: string } = {},
): XmlElement => {
  const paragraph = ctx.paragraph ?? null;
  const label = ctx.label ?? 'setShapeAnimation';
  const effect = oneOf(opts.effect, ANIMATION_EFFECTS, `${label}: effect`);
  const preset = PRESETS[effect];
  const family = familyOf(effect);
  const options = resolveAnimationOptions(effect, opts, label);
  // <p:cTn dur> is ST_TLTime (xsd:unsignedInt ms or "indefinite"). Bounds
  // checking rounds to whole milliseconds and rejects anything outside the
  // range, so we never emit an invalid dur.
  const duration =
    opts.durationMs === undefined
      ? family.durationMs
      : unsignedIntMs(opts.durationMs, `${label}: durationMs`);

  const start = oneOf(
    opts.start ?? 'click',
    ['click', 'withPrevious', 'afterPrevious'],
    `${label}: start`,
  );
  const delay = opts.delayMs === undefined ? 0 : unsignedIntMs(opts.delayMs, `${label}: delayMs`);

  const isEntrance = preset.presetClass === 'entr';
  const isExit = preset.presetClass === 'exit';

  // The behaviours, in the order PowerPoint writes them: an entrance reveals
  // the shape first and then moves it, an exit moves it and hides it one
  // millisecond before the end, and an emphasis effect never touches
  // visibility at all — the shape it turns is already on the slide.
  const motion: XmlElement[] = [];
  const firstMotionId = isEntrance ? 7 : 6;
  switch (family.motion) {
    case 'none':
      break;
    case 'fade':
      motion.push(
        buildAttrAnim({
          spid,
          id: firstMotionId,
          durationMs: duration,
          attrName: 'style.opacity',
          from: { flt: isEntrance ? '0' : '1' },
          to: { flt: isEntrance ? '1' : '0' },
          additive: true,
          paragraph,
        }),
      );
      break;
    case 'fly':
      motion.push(
        ...buildFlyAnims(spid, firstMotionId, duration, isEntrance, options.direction!, paragraph),
      );
      break;
    case 'peek':
      motion.push(
        buildPeekAnim(spid, firstMotionId, duration, isEntrance, options.direction!, paragraph),
        buildFilterEffect(
          spid,
          firstMotionId + 1,
          duration,
          isEntrance,
          family.filter!(options, isEntrance),
          paragraph,
        ),
      );
      break;
    case 'filter':
      motion.push(
        buildFilterEffect(
          spid,
          firstMotionId,
          duration,
          isEntrance,
          family.filter!(options, isEntrance),
          paragraph,
        ),
      );
      break;
    case 'zoom':
      motion.push(...buildZoomAnims(spid, firstMotionId, duration, isEntrance, paragraph));
      break;
    case 'spin':
      motion.push(buildSpinAnim(spid, firstMotionId, duration, paragraph));
      break;
  }

  const effectChildren: XmlElement[] = [];
  if (isEntrance) effectChildren.push(buildSetVisibility(spid, 6, true, 0, paragraph));
  effectChildren.push(...motion);
  if (isExit) {
    // An instant exit hides the shape straight away; one that animates has to
    // stay on the slide until it is over, so the kick trails the motion by the
    // one millisecond it takes itself.
    const hideAt = family.motion === 'none' ? 0 : trailingHideDelayMs(duration);
    effectChildren.push(buildSetVisibility(spid, 6 + motion.length, false, hideAt, paragraph));
  }

  // cTn id=5 — the effect node.
  const effectCTn = elem(NAME_C_TN, {
    attrs: [
      attr(ATTR_ID, '5'),
      attr(ATTR_PRESET_ID, String(family.presetId(options))),
      attr(ATTR_PRESET_CLASS, preset.presetClass),
      attr(ATTR_PRESET_SUBTYPE, String(preset.presetSubtype ?? family.presetSubtype(options))),
      attr(ATTR_FILL, 'hold'),
      attr(ATTR_GRP_ID, '0'),
      attr(ATTR_NODE_TYPE, START_NODE_TYPES[start]),
    ],
    children: [
      elem(NAME_ST_COND_LST, {
        children: [elem(NAME_COND, { attrs: [attr(ATTR_DELAY, String(delay))] })],
      }),
      elem(NAME_CHILD_TN_LST, { children: effectChildren }),
    ],
  });
  const effectPar = elem(NAME_PAR, { children: [effectCTn] });

  // cTn id=4 — the click wrapper.
  const clickWrapperCTn = elem(NAME_C_TN, {
    attrs: [attr(ATTR_ID, '4'), attr(ATTR_FILL, 'hold')],
    children: [
      elem(NAME_ST_COND_LST, {
        children: [elem(NAME_COND, { attrs: [attr(ATTR_DELAY, '0')] })],
      }),
      elem(NAME_CHILD_TN_LST, { children: [effectPar] }),
    ],
  });
  const clickWrapperPar = elem(NAME_PAR, { children: [clickWrapperCTn] });

  // cTn id=3 — the click stop. `indefinite` is what makes the group wait for
  // the viewer. A with/after effect that has no predecessor on the slide is
  // not waiting for anything, so its stop starts as the slide appears; when
  // the caller merges it behind an existing effect this whole wrapper is
  // discarded in favour of the one already there.
  const indefiniteCTn = elem(NAME_C_TN, {
    attrs: [attr(ATTR_ID, '3'), attr(ATTR_FILL, 'hold')],
    children: [
      elem(NAME_ST_COND_LST, {
        children: [
          elem(NAME_COND, { attrs: [attr(ATTR_DELAY, start === 'click' ? 'indefinite' : '0')] }),
        ],
      }),
      elem(NAME_CHILD_TN_LST, { children: [clickWrapperPar] }),
    ],
  });
  const indefinitePar = elem(NAME_PAR, { children: [indefiniteCTn] });

  // cTn id=2 — the mainSeq.
  const mainSeqCTn = elem(NAME_C_TN, {
    attrs: [attr(ATTR_ID, '2'), attr(ATTR_DUR, 'indefinite'), attr(ATTR_NODE_TYPE, 'mainSeq')],
    children: [elem(NAME_CHILD_TN_LST, { children: [indefinitePar] })],
  });

  // Slide-level next/prev navigation hooks.
  const prevCond = elem(NAME_PREV_COND_LST, {
    children: [
      elem(NAME_COND, {
        attrs: [attr(ATTR_EVT, 'onPrev'), attr(ATTR_DELAY, '0')],
        children: [elem(NAME_TGT_EL, { children: [elem(NAME_SLD_TGT)] })],
      }),
    ],
  });
  const nextCond = elem(NAME_NEXT_COND_LST, {
    children: [
      elem(NAME_COND, {
        attrs: [attr(ATTR_EVT, 'onNext'), attr(ATTR_DELAY, '0')],
        children: [elem(NAME_TGT_EL, { children: [elem(NAME_SLD_TGT)] })],
      }),
    ],
  });
  const seq = elem(NAME_SEQ, {
    attrs: [attr(ATTR_CONCURRENT, '1'), attr(ATTR_NEXT_AC, 'seek')],
    children: [mainSeqCTn, prevCond, nextCond],
  });

  // cTn id=1 — the tmRoot.
  const rootCTn = elem(NAME_C_TN, {
    attrs: [
      attr(ATTR_ID, '1'),
      attr(ATTR_DUR, 'indefinite'),
      attr(ATTR_RESTART, 'never'),
      attr(ATTR_NODE_TYPE, 'tmRoot'),
    ],
    children: [elem(NAME_CHILD_TN_LST, { children: [seq] })],
  });
  const rootPar = elem(NAME_PAR, { children: [rootCTn] });
  const tnLst = elem(NAME_TN_LST, { children: [rootPar] });

  // bldLst entry — required for PowerPoint to render the effect. `build="p"`
  // is what tells it the body is revealed paragraph by paragraph; the default
  // (`whole`) would reveal all of it on the first effect.
  const bldAttrs = [attr(ATTR_SPID, String(spid)), attr(ATTR_GRP_ID, '0')];
  if (paragraph !== null) bldAttrs.push(attr(ATTR_BUILD, 'p'));
  const bldLst = elem(NAME_BLD_LST, {
    children: [elem(NAME_BLD_P, { attrs: bldAttrs })],
  });

  return elem(NAME_TIMING, { children: [tnLst, bldLst] });
};
