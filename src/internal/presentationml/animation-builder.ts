// Animation builder — emits a `<p:timing>` block carrying a single effect on
// one target shape. The caller (`setShapeAnimation`) merges it into the slide's
// existing tree, so a slide ends up with as many effects as there were calls.
//
// Scope:
//
//   - Entrance, exit and one emphasis preset (`spin`). Motion paths
//     (`presetClass="path"`) and the rest of the emphasis family are not
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

/**
 * What kind of effect to apply.
 *
 * Entrance effects put the shape on the slide and exit effects take it off.
 * `spin` is the one emphasis effect modelled here: it turns a shape that is
 * already on the slide and leaves it exactly where it was, so it never decides
 * whether the shape is shown.
 */
export type AnimationEffect =
  | 'appear'
  | 'fadeIn'
  | 'flyIn'
  | 'zoomIn'
  | 'disappear'
  | 'fadeOut'
  | 'flyOut'
  | 'zoomOut'
  | 'spin';

/**
 * Which edge of the slide a `flyIn` comes from, or a `flyOut` leaves by —
 * PowerPoint's "From Bottom" / Google Slides' "Fly in from bottom".
 */
export type AnimationDirection = 'top' | 'right' | 'bottom' | 'left';

export const ANIMATION_EFFECTS: readonly AnimationEffect[] = [
  'appear',
  'fadeIn',
  'flyIn',
  'zoomIn',
  'disappear',
  'fadeOut',
  'flyOut',
  'zoomOut',
  'spin',
];

export const ANIMATION_DIRECTIONS: readonly AnimationDirection[] = [
  'top',
  'right',
  'bottom',
  'left',
];

/**
 * The `presetSubtype` PowerPoint writes for a directional preset. It is a
 * bitmask over the four edges, which is why the diagonals it also offers are
 * the two bits together (top-left is 1|8 = 9). Only the four straight
 * directions are modelled, so only the four single bits are written.
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
};

/** How an effect animates, beyond putting the shape on the slide or off it. */
type Motion = 'none' | 'fade' | 'fly' | 'zoom' | 'spin';

interface PresetDescriptor {
  readonly presetId: number;
  readonly presetClass: 'entr' | 'exit' | 'emph';
  /** `null` when the effect's direction supplies it. */
  readonly presetSubtype: number | null;
  readonly motion: Motion;
}

const PRESETS: Record<AnimationEffect, PresetDescriptor> = {
  appear: { presetId: 1, presetClass: 'entr', presetSubtype: 0, motion: 'none' },
  fadeIn: { presetId: 10, presetClass: 'entr', presetSubtype: 0, motion: 'fade' },
  flyIn: { presetId: 2, presetClass: 'entr', presetSubtype: null, motion: 'fly' },
  zoomIn: { presetId: 23, presetClass: 'entr', presetSubtype: 16, motion: 'zoom' },
  disappear: { presetId: 1, presetClass: 'exit', presetSubtype: 0, motion: 'none' },
  fadeOut: { presetId: 10, presetClass: 'exit', presetSubtype: 0, motion: 'fade' },
  flyOut: { presetId: 2, presetClass: 'exit', presetSubtype: null, motion: 'fly' },
  zoomOut: { presetId: 23, presetClass: 'exit', presetSubtype: 32, motion: 'zoom' },
  spin: { presetId: 8, presetClass: 'emph', presetSubtype: 0, motion: 'spin' },
};

/** Whether the effect flies, and so takes a direction. */
export const isDirectionalEffect = (effect: AnimationEffect): boolean =>
  PRESETS[effect].motion === 'fly';

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
   * Which edge of the slide a `'flyIn'` comes from, or a `'flyOut'` leaves by.
   * Defaults to `'bottom'`, PowerPoint's own default for the preset. Passing it
   * for an effect that does not fly is an error rather than a no-op: it would
   * otherwise read as a direction the file never records.
   */
  readonly direction?: AnimationDirection;
  /**
   * Animation length in milliseconds. Defaults to 500ms. `'appear'` and
   * `'disappear'` are instantaneous by definition of the preset and write no
   * timed behaviour at all, so it does not reach them.
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
  // Just outside the edge: the shape's centre one half-size beyond it.
  const offSlide = (axis: 'x' | 'y'): string => {
    const half = `${base(axis === 'x' ? 'w' : 'h')}/2`;
    const far = axis === 'x' ? direction === 'right' : direction === 'bottom';
    return far ? `1+${half}` : `0-${half}`;
  };
  const moves = (axis: 'x' | 'y'): boolean =>
    axis === 'x'
      ? direction === 'left' || direction === 'right'
      : direction === 'top' || direction === 'bottom';
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
  if (opts.direction !== undefined && preset.presetSubtype !== null) {
    throw new RangeError(
      `${label}: direction only applies to an effect that flies (${ANIMATION_EFFECTS.filter(
        isDirectionalEffect,
      ).join(', ')}); ${JSON.stringify(effect)} does not.`,
    );
  }
  const direction =
    preset.presetSubtype === null
      ? oneOf(opts.direction ?? 'bottom', ANIMATION_DIRECTIONS, `${label}: direction`)
      : null;
  // <p:cTn dur> is ST_TLTime (xsd:unsignedInt ms or "indefinite"). Bounds
  // checking rounds to whole milliseconds and rejects anything outside the
  // range, so we never emit an invalid dur.
  const duration =
    opts.durationMs === undefined ? 500 : unsignedIntMs(opts.durationMs, `${label}: durationMs`);

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
  switch (preset.motion) {
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
        ...buildFlyAnims(spid, firstMotionId, duration, isEntrance, direction!, paragraph),
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
    const hideAt = preset.motion === 'none' ? 0 : trailingHideDelayMs(duration);
    effectChildren.push(buildSetVisibility(spid, 6 + motion.length, false, hideAt, paragraph));
  }

  // cTn id=5 — the effect node.
  const effectCTn = elem(NAME_C_TN, {
    attrs: [
      attr(ATTR_ID, '5'),
      attr(ATTR_PRESET_ID, String(preset.presetId)),
      attr(ATTR_PRESET_CLASS, preset.presetClass),
      attr(ATTR_PRESET_SUBTYPE, String(preset.presetSubtype ?? DIRECTION_SUBTYPES[direction!])),
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
