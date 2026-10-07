// Animation builder — emits a `<p:timing>` block carrying a single effect on
// one target shape. The caller (`setShapeAnimation`) merges it into the slide's
// existing tree, so a slide ends up with as many effects as there were calls.
//
// Scope: every effect of PowerPoint's Entrance, Emphasis and Exit galleries,
// on the whole shape or on one paragraph of its text body (`<p:txEl><p:pRg>`).
// Motion paths (`presetClass="path"`) are not modelled.
//
// Preset numbers and behaviour shapes
// -----------------------------------
// ECMA-376 leaves `presetID` / `presetSubtype` as plain integers (§19.5.19,
// CT_TLCommonTimeNodeData) and `<p:attrName>` as a plain string: neither the
// preset catalogue nor the attribute vocabulary is in the standard. Both are
// PowerPoint conventions, so each effect is written exactly as PowerPoint
// writes it: the behaviours come from Mac PowerPoint 16.113's own output for
// each gallery entry (`animation-presets.generated.ts`, generated from the
// captures in `test/fixtures/native/animations/`). Only what an option decides
// — a fly's direction, a filter's pattern — and the length are filled in here.
//
// What the attribute names mean is likewise convention: `ppt_x` / `ppt_y` are
// the shape's centre as a fraction of the slide, `ppt_w` / `ppt_h` its size as
// a fraction of the slide, and `r` its rotation. Cross-checked against
// LibreOffice's OOXML import, which maps them to X / Y / Width / Height /
// Rotate (`oox/source/ppt/pptfilterhelpers.cxx`, `getAttributeConversionList`).

import { oneOf, unsignedIntMs } from '../bounds.ts';
import {
  NS,
  type XmlElement,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
} from '../xml/index.ts';
import {
  ANIMATION_PRESET_TEMPLATES,
  ANIMATION_TEMPLATE_ABBREVIATIONS,
  type AnimationPresetTemplate,
} from './animation-presets.generated.ts';

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
const NAME_SLD_TGT = qname('p', 'sldTgt', NS.pml);
const NAME_ATTR_NAME_LST = qname('p', 'attrNameLst', NS.pml);
const NAME_ATTR_NAME = qname('p', 'attrName', NS.pml);
const NAME_C_BHVR = qname('p', 'cBhvr', NS.pml);
const NAME_TO = qname('p', 'to', NS.pml);
const NAME_STR_VAL = qname('p', 'strVal', NS.pml);
const NAME_BLD_LST = qname('p', 'bldLst', NS.pml);
const NAME_BLD_P = qname('p', 'bldP', NS.pml);

const ATTR_ID = qname('', 'id', '');
const ATTR_DUR = qname('', 'dur', '');
const ATTR_RESTART = qname('', 'restart', '');
const ATTR_NODE_TYPE = qname('', 'nodeType', '');
const ATTR_FILL = qname('', 'fill', '');
const ATTR_DELAY = qname('', 'delay', '');
const ATTR_CONCURRENT = qname('', 'concurrent', '');
const ATTR_NEXT_AC = qname('', 'nextAc', '');
const ATTR_SPID = qname('', 'spid', '');
const ATTR_EVT = qname('', 'evt', '');
const ATTR_VAL = qname('', 'val', '');
const ATTR_GRP_ID = qname('', 'grpId', '');
const ATTR_BUILD = qname('', 'build', '');
const ATTR_ANIM_BG = qname('', 'animBg', '');
const ATTR_AUTO_REV = qname('', 'autoRev', '');

/**
 * What kind of effect to apply — one token per entry of PowerPoint's Entrance,
 * Emphasis and Exit galleries, named after the gallery entry.
 *
 * Entrance effects put the shape on the slide and end in `In` (`appear` aside);
 * exit effects take it off and end in `Out` (`disappear` aside). Emphasis
 * effects animate a shape that is already on the slide and never decide
 * whether it is shown.
 *
 * Where a gallery name would not tell two presets apart, the token does:
 * `zoomIn` is PowerPoint's "Zoom" (preset 53, which fades as it grows) and
 * `basicZoomIn` its "Basic Zoom" (preset 23); `floatIn` / `floatOut` are
 * "Float In" / "Float Out" (preset 42) and `floatingIn` / `floatingOut` the
 * "Float" entrance and exit (preset 30).
 */
export type AnimationEffect =
  // Entrance
  | 'appear'
  | 'blindsIn'
  | 'checkerboardIn'
  | 'dissolveIn'
  | 'flyIn'
  | 'peekIn'
  | 'randomBarsIn'
  | 'shapeIn'
  | 'splitIn'
  | 'stripsIn'
  | 'wedgeIn'
  | 'wheelIn'
  | 'wipeIn'
  | 'expandIn'
  | 'fadeIn'
  | 'swivelIn'
  | 'zoomIn'
  | 'centerRevolveIn'
  | 'floatIn'
  | 'growTurnIn'
  | 'riseUpIn'
  | 'spinnerIn'
  | 'basicZoomIn'
  | 'stretchIn'
  | 'boomerangIn'
  | 'bounceIn'
  | 'creditsIn'
  | 'curveUpIn'
  | 'dropIn'
  | 'flipIn'
  | 'floatingIn'
  | 'pinwheelIn'
  | 'spiralIn'
  | 'basicSwivelIn'
  | 'whipIn'
  // Exit
  | 'blindsOut'
  | 'checkerboardOut'
  | 'disappear'
  | 'dissolveOut'
  | 'flyOut'
  | 'peekOut'
  | 'randomBarsOut'
  | 'shapeOut'
  | 'splitOut'
  | 'stripsOut'
  | 'wedgeOut'
  | 'wheelOut'
  | 'wipeOut'
  | 'contractOut'
  | 'fadeOut'
  | 'swivelOut'
  | 'zoomOut'
  | 'centerRevolveOut'
  | 'collapseOut'
  | 'floatOut'
  | 'shrinkTurnOut'
  | 'sinkDownOut'
  | 'spinnerOut'
  | 'basicZoomOut'
  | 'stretchyOut'
  | 'boomerangOut'
  | 'bounceOut'
  | 'creditsOut'
  | 'curveDownOut'
  | 'dropOut'
  | 'flipOut'
  | 'floatingOut'
  | 'pinwheelOut'
  | 'spiralOut'
  | 'basicSwivelOut'
  | 'whipOut'
  // Emphasis
  | 'fillColor'
  | 'fontColor'
  | 'growShrink'
  | 'lineColor'
  | 'spin'
  | 'transparency'
  | 'boldFlash'
  | 'brushColor'
  | 'complementaryColor'
  | 'complementaryColor2'
  | 'contrastingColor'
  | 'darken'
  | 'desaturate'
  | 'lighten'
  | 'objectColor'
  | 'pulse'
  | 'underline'
  | 'colorPulse'
  | 'growWithColor'
  | 'shimmer'
  | 'teeter'
  | 'blink'
  | 'boldReveal'
  | 'wave';

// Every token has a template; test/fn-animation-native.test.ts checks that
// every template has a token.
const TEMPLATES: Readonly<Record<AnimationEffect, AnimationPresetTemplate>> =
  ANIMATION_PRESET_TEMPLATES;

/** Every effect, in gallery order: entrance, exit, emphasis. */
export const ANIMATION_EFFECTS = Object.keys(TEMPLATES) as readonly AnimationEffect[];

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

/**
 * How an effect on a shape with text treats the text — PowerPoint's Effect
 * Options "Sequence":
 *
 *   - `'asOneObject'` animates the shape and its text together, as one effect.
 *   - `'allAtOnce'` gives each paragraph an effect of its own, all starting
 *     together (`<p:bldP build="allAtOnce">`).
 *   - `'byParagraph'` gives each paragraph an effect of its own, one click (or
 *     start condition) apiece (`<p:bldP build="p">`).
 */
export type AnimationTextBuild = 'asOneObject' | 'allAtOnce' | 'byParagraph';

const SIDES: readonly AnimationDirection[] = ['top', 'right', 'bottom', 'left'];
const CORNERS: readonly AnimationDirection[] = ['topLeft', 'topRight', 'bottomLeft', 'bottomRight'];

export const ANIMATION_DIRECTIONS: readonly AnimationDirection[] = [...SIDES, ...CORNERS];

const ORIENTATIONS: readonly AnimationOrientation[] = ['horizontal', 'vertical'];
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

/**
 * An effect family whose Effect Options change what is written: the preset
 * numbers, and the `{…}` placeholders its template leaves for them.
 */
interface OptionFamily {
  /** The values each option may have; the first is PowerPoint's default. */
  readonly domains: AnimationOptionDomains;
  /** The same for the exit, where PowerPoint's default differs. */
  readonly exitDomains?: AnimationOptionDomains;
  readonly presetId?: (o: AnimationEffectOptions) => number;
  readonly presetSubtype: (o: AnimationEffectOptions) => number;
  /** The placeholder values for an entrance (`entering`) or an exit. */
  readonly values: (o: AnimationEffectOptions, entering: boolean) => Record<string, string>;
}

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
 * The keyframes of a fly. Both axes are written even though only one of them
 * may move: PowerPoint holds the other one steady.
 *
 * `#` marks the shape's own authored value. Entrance keyframes carry it and
 * exit keyframes do not, which is the difference between the two preset rows.
 * Off the slide is the shape's centre one half-size beyond the edge.
 */
const flyValues = (direction: AnimationDirection, entering: boolean): Record<string, string> => {
  const base = (axis: 'x' | 'y' | 'w' | 'h'): string => `${entering ? '#' : ''}ppt_${axis}`;
  const edge = { x: horizontalEdge(direction), y: verticalEdge(direction) };
  const out: Record<string, string> = {};
  for (const axis of ['x', 'y'] as const) {
    const own = base(axis);
    const half = `${base(axis === 'x' ? 'w' : 'h')}/2`;
    const far = edge[axis] === 'right' || edge[axis] === 'bottom';
    const away = edge[axis] === null ? own : far ? `1+${half}` : `0-${half}`;
    const key = axis.toUpperCase();
    out[`${key}0`] = entering ? away : own;
    out[`${key}1`] = entering ? own : away;
  }
  return out;
};

const filterOnly =
  (filter: (o: AnimationEffectOptions, entering: boolean) => string) =>
  (o: AnimationEffectOptions, entering: boolean): Record<string, string> => ({
    F: filter(o, entering),
  });

const EDGES_FROM_BOTTOM: readonly AnimationDirection[] = [
  'bottom',
  ...SIDES.filter((d) => d !== 'bottom'),
];

// What PowerPoint writes for each Effect Options choice — preset id, subtype
// and filter — as its object model records them (MsoAnimEffect, PowerPoint 16;
// see POWERPOINT_PARITY.md "Animations"). The defaults are the ones the
// captures were made with, and the native-capture test checks them.
const OPTION_FAMILIES: Readonly<Record<string, OptionFamily>> = {
  fly: {
    domains: { direction: [...EDGES_FROM_BOTTOM, ...CORNERS] },
    presetSubtype: (o) => DIRECTION_SUBTYPES[o.direction!],
    values: (o, entering) => flyValues(o.direction!, entering),
  },
  peek: {
    domains: { direction: EDGES_FROM_BOTTOM },
    presetSubtype: (o) => DIRECTION_SUBTYPES[o.direction!],
    // The object slides 1.125 of its own size from just past the edge while
    // the wipe uncovers it from the far side, so an entrance wipes towards the
    // opposite edge. Unlike a fly, only the moving axis is written, and the
    // `#` stays on both ends of an exit too.
    values: (o, entering) => {
      const direction = o.direction!;
      const horizontal = horizontalEdge(direction) !== null;
      const axis = horizontal ? 'x' : 'y';
      const size = horizontal ? 'w' : 'h';
      const sign = direction === 'right' || direction === 'bottom' ? '+' : '-';
      const own = `#ppt_${axis}`;
      const away = `#ppt_${axis}${sign}#ppt_${size}*1.125000`;
      return {
        A: `ppt_${axis}`,
        P0: entering ? away : own,
        P1: entering ? own : away,
        F: `wipe(${edgeWord[entering ? oppositeEdge[direction]! : direction]})`,
      };
    },
  },
  wipe: {
    domains: { direction: EDGES_FROM_BOTTOM },
    presetSubtype: (o) => DIRECTION_SUBTYPES[o.direction!],
    values: filterOnly((o) => `wipe(${edgeWord[o.direction!]})`),
  },
  blinds: {
    domains: { orientation: ORIENTATIONS },
    presetSubtype: (o) => ORIENTATION_SUBTYPES[o.orientation!],
    values: filterOnly((o) => `blinds(${o.orientation})`),
  },
  checkerboard: {
    domains: { orientation: ORIENTATIONS },
    presetSubtype: (o) => ORIENTATION_SUBTYPES[o.orientation!],
    values: filterOnly((o) => `checkerboard(${o.orientation === 'horizontal' ? 'across' : 'down'})`),
  },
  randomBars: {
    domains: { orientation: ORIENTATIONS },
    presetSubtype: (o) => ORIENTATION_SUBTYPES[o.orientation!],
    // ECMA-376's filter table spells this `randomBars(…)`; PowerPoint writes
    // and reads `randombar(…)`, so that is what is written.
    values: filterOnly((o) => `randombar(${o.orientation})`),
  },
  shape: {
    domains: { shape: SHAPES, inOut: ['in', 'out'] },
    // PowerPoint's Shape exit closes onto the centre rather than opening.
    exitDomains: { shape: SHAPES, inOut: ['out', 'in'] },
    presetId: (o) => SHAPE_PRESET_IDS[o.shape!],
    presetSubtype: (o) => IN_OUT_SUBTYPES[o.inOut!],
    values: filterOnly((o) => `${o.shape}(${o.inOut})`),
  },
  split: {
    domains: { orientation: ['vertical', 'horizontal'], inOut: ['in', 'out'] },
    presetSubtype: (o) =>
      IN_OUT_SUBTYPES[o.inOut!] +
      (o.orientation === 'vertical' ? 5 : ORIENTATION_SUBTYPES.horizontal),
    values: filterOnly((o) => `barn(${o.inOut}${capitalised(o.orientation!)})`),
  },
  strips: {
    domains: { direction: ['bottomLeft', 'topLeft', 'topRight', 'bottomRight'] },
    presetSubtype: (o) => DIRECTION_SUBTYPES[o.direction!],
    values: filterOnly((o) => `strips(${cornerWord[o.direction!]})`),
  },
  wheel: {
    domains: { spokes: WHEEL_SPOKES },
    presetSubtype: (o) => o.spokes!,
    values: filterOnly((o) => `wheel(${o.spokes})`),
  },
};

const NO_OPTIONS: AnimationOptionDomains = {};

const optionFamilyOf = (effect: AnimationEffect): OptionFamily | undefined =>
  OPTION_FAMILIES[effect.replace(/(In|Out)$/, '')];

const presetClassOf = (effect: AnimationEffect): 'entr' | 'exit' | 'emph' => TEMPLATES[effect][0];

/** The options `effect` takes and the values each may have, the default first. */
export const animationOptionDomains = (effect: AnimationEffect): AnimationOptionDomains => {
  const family = optionFamilyOf(effect);
  if (family === undefined) return NO_OPTIONS;
  return presetClassOf(effect) === 'exit' ? (family.exitDomains ?? family.domains) : family.domains;
};

/** Whether the effect flies, and so takes any of the eight directions. */
export const isDirectionalEffect = (effect: AnimationEffect): boolean =>
  effect === 'flyIn' || effect === 'flyOut';

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
  const domains = animationOptionDomains(effect);
  const pick = <T extends string | number>(
    name: AnimationOptionName,
    domain: readonly T[] | undefined,
  ): T | null => {
    const value = given[name];
    if (domain === undefined) {
      if (value === undefined) return null;
      throw new RangeError(
        `${label}: ${name} only applies to ${ANIMATION_EFFECTS.filter(
          (other) => animationOptionDomains(other)[name] !== undefined,
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

const presetIdOf = (effect: AnimationEffect, options: AnimationEffectOptions): number =>
  optionFamilyOf(effect)?.presetId?.(options) ?? TEMPLATES[effect][1];

const presetSubtypeOf = (effect: AnimationEffect, options: AnimationEffectOptions): number =>
  optionFamilyOf(effect)?.presetSubtype(options) ?? TEMPLATES[effect][2];

/**
 * Which `<p:bldP>` PowerPoint writes for the effect: `'shape'` animates the
 * shape's background along with its text (`animBg`), `'text'` — the effects
 * that only restyle text, such as Font Color or Underline — the text alone,
 * and `'none'` — Fill Color and Line Color — writes no build entry and leaves
 * the effect out of every build group.
 */
export const animationBuildKind = (effect: AnimationEffect): 'shape' | 'text' | 'none' =>
  TEMPLATES[effect][4];

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
    const domains = animationOptionDomains(effect);
    return combinations(domains).map((options) => ({
      presetClass: presetClassOf(effect),
      presetId: presetIdOf(effect, options),
      presetSubtype: presetSubtypeOf(effect, options),
      effect,
      options,
      optionless: domains === NO_OPTIONS,
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

// ---------------------------------------------------------------------------
// An effect's length.
//
// PowerPoint's Duration is when the effect's last behaviour ends: a behaviour
// that reverses (`autoRev`) runs twice, and one that waits starts that much
// later. Changing it scales every behaviour's length and wait by the same
// factor, which is how Bounce keeps its shape at any speed. The writer and the
// editing path share these, so a fresh effect and a retimed one agree.

const isPml = (el: XmlElement, local: string): boolean =>
  el.name.namespaceURI === NS.pml && el.name.localName === local;

const BEHAVIOUR_LOCALS = new Set([
  'set',
  'anim',
  'animClr',
  'animEffect',
  'animMotion',
  'animRot',
  'animScale',
  'cmd',
]);

/** A behaviour of an effect, and the `<p:cTn>` that times it. */
interface Behaviour {
  readonly element: XmlElement;
  readonly cTn: XmlElement | null;
  /** The 1ms `<p:set>` that puts the target on the slide or takes it off. */
  readonly kick: boolean;
}

const attrNameOf = (cBhvr: XmlElement): string | null => {
  const list = firstChildElement(cBhvr, NAME_ATTR_NAME_LST);
  const name = list === null ? null : firstChildElement(list, NAME_ATTR_NAME);
  return name?.children.find((c) => c.kind === 'text')?.data ?? null;
};

/**
 * The behaviours directly under an effect node, or `null` when the node nests
 * time containers of its own — whose children start when *they* do, so the
 * effect's clock no longer says when they run.
 */
export const effectBehaviours = (effectCTn: XmlElement): Behaviour[] | null => {
  const list = firstChildElement(effectCTn, NAME_CHILD_TN_LST);
  if (list === null) return [];
  const out: Behaviour[] = [];
  for (const child of list.children) {
    if (child.kind !== 'element') continue;
    if (child.name.namespaceURI !== NS.pml || !BEHAVIOUR_LOCALS.has(child.name.localName)) {
      return null;
    }
    const cBhvr = firstChildElement(child, NAME_C_BHVR);
    const cTn = cBhvr === null ? null : firstChildElement(cBhvr, NAME_C_TN);
    out.push({
      element: child,
      cTn,
      kick: isPml(child, 'set') && cBhvr !== null && attrNameOf(cBhvr) === 'style.visibility',
    });
  }
  return out;
};

const wholeMs = (raw: string | null): number | null => {
  if (raw === null || !/^\d+$/.test(raw)) return null;
  const n = Number(raw);
  return Number.isSafeInteger(n) ? n : null;
};

/** The single plain-offset `<p:cond>` a behaviour starts on; `undefined` when it has none. */
const offsetCondOf = (cTn: XmlElement): XmlElement | null | undefined => {
  const stCondLst = firstChildElement(cTn, NAME_ST_COND_LST);
  if (stCondLst === null) return undefined;
  const conds = stCondLst.children.filter(
    (c): c is XmlElement => c.kind === 'element' && isPml(c, 'cond'),
  );
  if (conds.length !== 1) return null;
  const cond = conds[0]!;
  if (getAttrValue(cond, ATTR_EVT) !== null) return null;
  if (cond.children.some((c) => c.kind === 'element')) return null;
  return cond;
};

/** Attributes that repeat or rescale a node by more than reversing it once. */
const REPEATING_ATTRS = new Set(['repeatCount', 'repeatDur', 'spd']);

const reverses = (cTn: XmlElement): boolean => {
  const raw = getAttrValue(cTn, ATTR_AUTO_REV);
  return raw === '1' || raw === 'true';
};

/** When one behaviour ends, measured from the start of its effect; `null` when unstated. */
const behaviourEndMs = (cTn: XmlElement): number | null => {
  if (cTn.attrs.some((a) => a.name.namespaceURI === '' && REPEATING_ATTRS.has(a.name.localName))) {
    return null;
  }
  const dur = wholeMs(getAttrValue(cTn, ATTR_DUR));
  const cond = offsetCondOf(cTn);
  const delay = cond === undefined ? 0 : cond === null ? null : wholeMs(getAttrValue(cond, ATTR_DELAY));
  if (dur === null || delay === null) return null;
  return delay + dur * (reverses(cTn) ? 2 : 1);
};

/**
 * PowerPoint's Duration for an effect: when the last of its behaviours ends,
 * the visibility kicks aside. `null` when any behaviour runs indefinitely,
 * states no length, repeats or waits on something other than a plain offset,
 * and when there is no timed behaviour at all (`appear`).
 */
export const effectDurationMs = (effectCTn: XmlElement): number | null =>
  lastBehaviourEndMs(effectCTn, false);

/**
 * When the last of an effect's behaviours ends, the visibility kicks included —
 * for `appear` the kick is the whole effect. `null` when any behaviour cannot
 * be placed on the effect's clock.
 */
export const lastBehaviourEndMs = (effectCTn: XmlElement, withKicks: boolean): number | null => {
  const behaviours = effectBehaviours(effectCTn);
  if (behaviours === null) return null;
  let longest: number | null = null;
  for (const behaviour of behaviours) {
    if (behaviour.kick && !withKicks) continue;
    const end = behaviour.cTn === null ? null : behaviourEndMs(behaviour.cTn);
    if (end === null) return null;
    longest = Math.max(longest ?? 0, end);
  }
  return longest;
};

const setAttrValue = (el: XmlElement, name: string, value: string): void => {
  const has = el.attrs.some((a) => a.name.namespaceURI === '' && a.name.localName === name);
  el.attrs = has
    ? el.attrs.map((a) =>
        a.name.namespaceURI === '' && a.name.localName === name ? { ...a, value } : a,
      )
    : [...el.attrs, attr(qname('', name, ''), value)];
};

const hidesTarget = (behaviour: Behaviour): boolean => {
  const to = firstChildElement(behaviour.element, NAME_TO);
  const strVal = to === null ? null : firstChildElement(to, NAME_STR_VAL);
  return strVal !== null && getAttrValue(strVal, ATTR_VAL) === 'hidden';
};

/**
 * Retimes an effect to run for `durationMs`, in place: every behaviour's length
 * and wait scale by the same factor. The `<p:set>` that ends an exit trails
 * the motion by its own millisecond, so it moves to the new end rather than
 * scaling — but only when it stands at the old end: a kick anywhere else was
 * put there by timing this library did not write.
 *
 * `false` when the effect has no length to scale from (`effectDurationMs` is
 * `null`). An effect with no timed behaviour — `appear`, `disappear` —
 * has nothing to retime and is left as it is.
 */
export const setEffectDurationMs = (effectCTn: XmlElement, durationMs: number): boolean => {
  const behaviours = effectBehaviours(effectCTn);
  if (behaviours === null) return false;
  const timed = behaviours.filter((b) => !b.kick);
  if (timed.length === 0) return true;
  const was = effectDurationMs(effectCTn);
  if (was === null) return false;
  // An effect shortened to nothing has no proportions left to keep, so every
  // behaviour runs the whole new length.
  const scale = (ms: number): string =>
    String(was === 0 ? durationMs : Math.round((ms * durationMs) / was));
  for (const { cTn } of timed) {
    setAttrValue(cTn!, 'dur', scale(wholeMs(getAttrValue(cTn!, ATTR_DUR))!));
    const cond = offsetCondOf(cTn!);
    if (cond) setAttrValue(cond, 'delay', scale(wholeMs(getAttrValue(cond, ATTR_DELAY))!));
  }
  const trailing = String(trailingHideDelayMs(was));
  for (const kick of behaviours) {
    if (!kick.kick || kick.cTn === null || !hidesTarget(kick)) continue;
    const cond = offsetCondOf(kick.cTn);
    if (!cond || getAttrValue(cond, ATTR_DELAY) !== trailing) continue;
    setAttrValue(cond, 'delay', String(trailingHideDelayMs(durationMs)));
  }
  return true;
};

// ---------------------------------------------------------------------------
// Templates.

/**
 * A template body back as markup: the abbreviations undone in the reverse of
 * the order they were applied, then the keyframe, attribute-name and delay
 * shorthands, then the `p:` prefix on every element that has none.
 */
const expandTemplate = (body: string): string => {
  let out = body;
  for (let at = ANIMATION_TEMPLATE_ABBREVIATIONS.length - 1; at >= 0; at -= 1) {
    out = out.replaceAll(`~${String.fromCharCode(97 + at)}`, ANIMATION_TEMPLATE_ABBREVIATIONS[at]!);
  }
  return out
    .replaceAll(/\^(\w+)\^/g, '<stCondLst><cond delay="$1"/></stCondLst>')
    .replaceAll(/@([^@]*)@/g, '<attrNameLst><attrName>$1</attrName></attrNameLst>')
    .replaceAll(/\[(\d+)!([^\]]*)\]/g, '<tav tm="$1"><val><fltVal val="$2"/></val></tav>')
    .replaceAll(/\[(\d+)\|([^\]]*)\]/g, '<tav tm="$1"><val><strVal val="$2"/></val></tav>')
    .replaceAll(/<(\/?)([A-Za-z]+)(?=[\s/>])/g, '<$1p:$2');
};

/** The `<p:tgtEl>` markup for a shape, or one paragraph of it (CT_IndexRange is inclusive). */
const targetMarkup = (spid: number, paragraph: number | null): string =>
  paragraph === null
    ? `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>`
    : `<p:tgtEl><p:spTgt spid="${spid}"><p:txEl><p:pRg st="${paragraph}" end="${paragraph}"/></p:txEl></p:spTgt></p:tgtEl>`;

/** When an effect runs relative to the one before it. */
export type AnimationStartCondition = 'click' | 'withPrevious' | 'afterPrevious';

// ST_TLTimeNodeType tokens for the three start conditions.
const START_NODE_TYPES: Record<AnimationStartCondition, string> = {
  click: 'clickEffect',
  withPrevious: 'withEffect',
  afterPrevious: 'afterEffect',
};

/**
 * The effect node PowerPoint writes for `effect` at its own default length,
 * with the given target, start and delay. Its cTn ids are left for the caller
 * to number.
 */
const effectNode = (
  effect: AnimationEffect,
  options: AnimationEffectOptions,
  target: string,
  start: AnimationStartCondition,
  delayMs: number,
): XmlElement => {
  const [presetClass, , , attrs, build, body] = TEMPLATES[effect];
  const values = optionFamilyOf(effect)?.values(options, presetClass === 'entr') ?? {};
  const inner = expandTemplate(body).replaceAll(/\{(\w+)\}/g, (_, name: string) =>
    name === 'T' ? target : values[name]!,
  );
  const head = [
    `presetID="${presetIdOf(effect, options)}"`,
    `presetClass="${presetClass}"`,
    `presetSubtype="${presetSubtypeOf(effect, options)}"`,
    attrs,
    build === 'none' ? '' : 'grpId="0"',
    `nodeType="${START_NODE_TYPES[start]}"`,
  ]
    .filter((part) => part !== '')
    .join(' ');
  const xml =
    `<p:cTn xmlns:p="${NS.pml}" xmlns:a="${NS.dml}" ${head}>` +
    `<p:stCondLst><p:cond delay="${delayMs}"/></p:stCondLst>${inner}</p:cTn>`;
  const cTn = parseXml(xml).root;
  // The slide declares both prefixes; the copy here only had to parse.
  cTn.prefixDecls = new Map();
  return cTn;
};

// Worked out from the template the first time it is asked for, rather than for
// all 95 effects when the module loads.
const defaultDurations = new Map<AnimationEffect, number | null>();

/** PowerPoint's default length for the preset in milliseconds; `null` when it holds until the slide ends. */
export const defaultAnimationDurationMs = (effect: AnimationEffect): number | null => {
  if (!defaultDurations.has(effect)) {
    const options = resolveAnimationOptions(effect, {}, 'defaultAnimationDurationMs');
    const node = effectNode(effect, options, targetMarkup(1, null), 'click', 0);
    defaultDurations.set(effect, effectDurationMs(node));
  }
  return defaultDurations.get(effect)!;
};

/** The local names of the behaviour elements PowerPoint writes for `effect`. */
export const effectBehaviourNames = (effect: AnimationEffect): ReadonlySet<string> => {
  const names = new Set<string>();
  for (const match of expandTemplate(TEMPLATES[effect][5]).matchAll(/<p:(\w+)/g)) {
    if (BEHAVIOUR_LOCALS.has(match[1]!)) names.add(match[1]!);
  }
  return names;
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
  /**
   * For shape and split: open from the centre (`'out'`) or close in on it
   * (`'in'`). Defaults to `'in'` — `'out'` for `'shapeOut'`, as in PowerPoint.
   */
  readonly inOut?: AnimationInOut;
  /** For `'shapeIn'` / `'shapeOut'`: the outline. Defaults to `'circle'`. */
  readonly shape?: AnimationShape;
  /** For `'wheelIn'` / `'wheelOut'`: 1, 2, 3, 4 or 8 spokes. Defaults to 1. */
  readonly spokes?: number;
  /**
   * How long the effect runs, in milliseconds. Defaults to PowerPoint's default
   * for the preset (`defaultAnimationDurationMs`). An effect made of several
   * behaviours is scaled as a whole, the way PowerPoint's Duration box does.
   * `'appear'` and `'disappear'` are instantaneous and write no timed
   * behaviour, so it does not reach them; `'transparency'` and `'boldReveal'`
   * hold until the end of the slide and take none.
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
   * How the shape's text is animated — PowerPoint's Effect Options "Sequence".
   * Defaults to `'asOneObject'`, the gallery's own default. `'allAtOnce'` and
   * `'byParagraph'` give every paragraph an effect of its own: started together
   * for the first, each on its own start (a click apiece by default) for the
   * second. Fill Color and Line Color animate the shape and take neither.
   */
  readonly build?: AnimationTextBuild;
}

/**
 * Numbers every `<p:cTn>` under `el` in document order from `first`, the way
 * PowerPoint numbers a tree it writes.
 */
const numberCTns = (el: XmlElement, first: number): number => {
  let next = first;
  const walk = (e: XmlElement): void => {
    if (isPml(e, 'cTn')) {
      const id = String(next++);
      const has = e.attrs.some((a) => a.name.namespaceURI === '' && a.name.localName === 'id');
      e.attrs = has
        ? e.attrs.map((a) =>
            a.name.namespaceURI === '' && a.name.localName === 'id' ? { ...a, value: id } : a,
          )
        : [attr(ATTR_ID, id), ...e.attrs];
    }
    for (const c of e.children) if (c.kind === 'element') walk(c);
  };
  walk(el);
  return next;
};

/** What the caller knows about the target that the effect's own options do not say. */
export interface EffectContext {
  /** The paragraph the effect animates; `null` (the default) for the whole shape. */
  readonly paragraph?: number | null;
  /** The build the paragraph belongs to; ignored for the whole shape. */
  readonly build?: AnimationTextBuild;
  /**
   * Whether the shape draws a background — a fill or an outline — that animates
   * with its text. PowerPoint marks the build entry `animBg` then.
   */
  readonly background?: boolean;
  /** Overrides `opts.start` — the later paragraphs of an all-at-once build run with the first. */
  readonly start?: AnimationStartCondition;
  readonly label?: string;
}

/**
 * Builds the complete `<p:timing>` element for a single effect on the given
 * shape id. The result is a standalone tree whose outermost `<p:par>` is one
 * click stop; merging it behind effects that already exist is the caller's
 * job, and only the caller knows whether that stop survives.
 *
 * A paragraph build is one such effect per paragraph, which is why the caller
 * loops rather than this function: the effects share a build group and each
 * needs merging into whatever the slide holds by then.
 */
export const buildSingleEffectTiming = (
  spid: number,
  opts: AnimationOptions,
  ctx: EffectContext = {},
): XmlElement => {
  const paragraph = ctx.paragraph ?? null;
  const label = ctx.label ?? 'setShapeAnimation';
  const effect = oneOf(opts.effect, ANIMATION_EFFECTS, `${label}: effect`);
  const options = resolveAnimationOptions(effect, opts, label);
  const start = oneOf(
    ctx.start ?? opts.start ?? 'click',
    ['click', 'withPrevious', 'afterPrevious'],
    `${label}: start`,
  );
  const delay = opts.delayMs === undefined ? 0 : unsignedIntMs(opts.delayMs, `${label}: delayMs`);

  const effectCTn = effectNode(effect, options, targetMarkup(spid, paragraph), start, delay);
  // <p:cTn dur> is ST_TLTime (xsd:unsignedInt ms or "indefinite"). Bounds
  // checking rounds to whole milliseconds and rejects anything outside the
  // range, so we never emit an invalid dur.
  if (
    opts.durationMs !== undefined &&
    !setEffectDurationMs(effectCTn, unsignedIntMs(opts.durationMs, `${label}: durationMs`))
  ) {
    throw new RangeError(
      `${label}: ${JSON.stringify(effect)} holds until the end of the slide, so it takes no durationMs.`,
    );
  }
  const effectPar = elem(NAME_PAR, { children: [effectCTn] });

  // The click wrapper.
  const clickWrapperCTn = elem(NAME_C_TN, {
    attrs: [attr(ATTR_FILL, 'hold')],
    children: [
      elem(NAME_ST_COND_LST, {
        children: [elem(NAME_COND, { attrs: [attr(ATTR_DELAY, '0')] })],
      }),
      elem(NAME_CHILD_TN_LST, { children: [effectPar] }),
    ],
  });
  const clickWrapperPar = elem(NAME_PAR, { children: [clickWrapperCTn] });

  // The click stop. `indefinite` is what makes the group wait for the viewer.
  // A with/after effect that has no predecessor on the slide is not waiting
  // for anything, so its stop starts as the slide appears; when the caller
  // merges it behind an existing effect this whole wrapper is discarded in
  // favour of the one already there.
  const indefiniteCTn = elem(NAME_C_TN, {
    attrs: [attr(ATTR_FILL, 'hold')],
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

  const mainSeqCTn = elem(NAME_C_TN, {
    attrs: [attr(ATTR_DUR, 'indefinite'), attr(ATTR_NODE_TYPE, 'mainSeq')],
    children: [elem(NAME_CHILD_TN_LST, { children: [indefinitePar] })],
  });

  // Slide-level next/prev navigation hooks.
  const navigation = (list: typeof NAME_PREV_COND_LST, evt: string): XmlElement =>
    elem(list, {
      children: [
        elem(NAME_COND, {
          attrs: [attr(ATTR_EVT, evt), attr(ATTR_DELAY, '0')],
          children: [elem(NAME_TGT_EL, { children: [elem(NAME_SLD_TGT)] })],
        }),
      ],
    });
  const seq = elem(NAME_SEQ, {
    attrs: [attr(ATTR_CONCURRENT, '1'), attr(ATTR_NEXT_AC, 'seek')],
    children: [
      mainSeqCTn,
      navigation(NAME_PREV_COND_LST, 'onPrev'),
      navigation(NAME_NEXT_COND_LST, 'onNext'),
    ],
  });

  const rootCTn = elem(NAME_C_TN, {
    attrs: [
      attr(ATTR_DUR, 'indefinite'),
      attr(ATTR_RESTART, 'never'),
      attr(ATTR_NODE_TYPE, 'tmRoot'),
    ],
    children: [elem(NAME_CHILD_TN_LST, { children: [seq] })],
  });
  const tnLst = elem(NAME_TN_LST, { children: [elem(NAME_PAR, { children: [rootCTn] })] });
  numberCTns(tnLst, 1);

  // The build entry PowerPoint needs to render the effect. `build` tells it
  // the body is revealed paragraph by paragraph; without it (`whole`) the
  // first paragraph's effect would reveal all of them. `animBg` animates the
  // shape's own fill and outline with its text.
  const kind = animationBuildKind(effect);
  if (kind === 'none') return elem(NAME_TIMING, { children: [tnLst] });
  const bldAttrs = [attr(ATTR_SPID, String(spid)), attr(ATTR_GRP_ID, '0')];
  if (paragraph !== null) {
    bldAttrs.push(attr(ATTR_BUILD, ctx.build === 'allAtOnce' ? 'allAtOnce' : 'p'));
  } else if (kind === 'shape' && ctx.background === true) {
    bldAttrs.push(attr(ATTR_ANIM_BG, '1'));
  }
  const bldLst = elem(NAME_BLD_LST, { children: [elem(NAME_BLD_P, { attrs: bldAttrs })] });
  return elem(NAME_TIMING, { children: [tnLst, bldLst] });
};
