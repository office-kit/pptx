// Animation builder — emits a `<p:timing>` block carrying a single
// click-triggered entrance / exit effect on one target shape.
//
// Scope (v1):
//
//   - Exactly one effect per slide. Calling `setShapeAnimation` on a
//     second shape replaces the first. The plan calls this out as the
//     curated subset; full timing-tree authoring is post-1.0.
//   - Click trigger only. After-previous / with-previous chaining lands
//     when multi-effect support does.
//   - Entrance + exit preset families. Emphasis presets are post-1.0.
//
// The timing tree shape follows what PowerPoint itself emits for a
// single "fade in on click" effect — the boilerplate scaffolding around
// the actual `<p:set>` / `<p:anim>` is fixed; we just swap presetID,
// presetClass, and the target spid.

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
const NAME_TAV_LST = qname('p', 'tavLst', NS.pml);
const NAME_TAV = qname('p', 'tav', NS.pml);
const NAME_BLD_LST = qname('p', 'bldLst', NS.pml);
const NAME_BLD_P = qname('p', 'bldP', NS.pml);

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
const ATTR_TM = qname('', 'tm', '');

/** What kind of effect to apply. v1 covers the four most-used presets. */
export type AnimationEffect = 'fadeIn' | 'fadeOut' | 'appear' | 'disappear';

interface PresetDescriptor {
  readonly presetId: number;
  readonly presetClass: 'entr' | 'exit';
  readonly presetSubtype: number;
}

const PRESETS: Record<AnimationEffect, PresetDescriptor> = {
  appear: { presetId: 1, presetClass: 'entr', presetSubtype: 0 },
  fadeIn: { presetId: 10, presetClass: 'entr', presetSubtype: 0 },
  disappear: { presetId: 1, presetClass: 'exit', presetSubtype: 0 },
  fadeOut: { presetId: 10, presetClass: 'exit', presetSubtype: 0 },
};

export interface AnimationOptions {
  /** Which preset effect to apply. */
  readonly effect: AnimationEffect;
  /** Animation length in milliseconds. Defaults to 500ms. */
  readonly durationMs?: number;
}

const buildSetVisibility = (spid: number, visible: boolean): XmlElement => {
  const tgt = elem(NAME_TGT_EL, {
    children: [elem(NAME_SP_TGT, { attrs: [attr(ATTR_SPID, String(spid))] })],
  });
  const attrName = elem(NAME_ATTR_NAME_LST, {
    children: [elem(NAME_ATTR_NAME_FN, { children: [{ kind: 'text', data: 'style.visibility' }] })],
  });
  const cTn = elem(NAME_C_TN, {
    attrs: [attr(ATTR_ID, '6'), attr(ATTR_DUR, '1'), attr(ATTR_FILL, 'hold')],
    children: [
      elem(NAME_ST_COND_LST, {
        children: [elem(NAME_COND, { attrs: [attr(ATTR_DELAY, '0')] })],
      }),
    ],
  });
  const cBhvr = elem(NAME_C_BHVR, { children: [cTn, tgt, attrName] });
  const to = elem(NAME_TO, {
    children: [elem(NAME_STR_VAL, { attrs: [attr(ATTR_VAL, visible ? 'visible' : 'hidden')] })],
  });
  return elem(NAME_SET, { children: [cBhvr, to] });
};

const buildOpacityAnim = (spid: number, durationMs: number, fadeIn: boolean): XmlElement => {
  const tgt = elem(NAME_TGT_EL, {
    children: [elem(NAME_SP_TGT, { attrs: [attr(ATTR_SPID, String(spid))] })],
  });
  const attrName = elem(NAME_ATTR_NAME_LST, {
    children: [elem(NAME_ATTR_NAME_FN, { children: [{ kind: 'text', data: 'style.opacity' }] })],
  });
  const cTn = elem(NAME_C_TN, {
    attrs: [attr(ATTR_ID, '7'), attr(ATTR_DUR, String(durationMs)), attr(ATTR_FILL, 'hold')],
  });
  const cBhvr = elem(NAME_C_BHVR, {
    attrs: [attr(ATTR_ADDITIVE, 'base')],
    children: [cTn, tgt, attrName],
  });
  const fromVal = fadeIn ? '0' : '1';
  const toVal = fadeIn ? '1' : '0';
  const tavLst = elem(NAME_TAV_LST, {
    children: [
      elem(NAME_TAV, {
        attrs: [attr(ATTR_TM, '0')],
        children: [
          elem(NAME_VAL, {
            children: [elem(NAME_FLT_VAL, { attrs: [attr(ATTR_VAL, fromVal)] })],
          }),
        ],
      }),
      elem(NAME_TAV, {
        attrs: [attr(ATTR_TM, '100000')],
        children: [
          elem(NAME_VAL, {
            children: [elem(NAME_FLT_VAL, { attrs: [attr(ATTR_VAL, toVal)] })],
          }),
        ],
      }),
    ],
  });
  return elem(NAME_ANIM, {
    attrs: [attr(ATTR_CALCMODE, 'lin'), attr(ATTR_VALUE_TYPE, 'num')],
    children: [cBhvr, tavLst],
  });
};

/**
 * Builds the complete `<p:timing>` element for a single click-effect on
 * the given shape id. Returns null for unsupported effect kinds.
 */
export const buildSingleEffectTiming = (spid: number, opts: AnimationOptions): XmlElement => {
  const preset = PRESETS[opts.effect];
  const duration = opts.durationMs ?? 500;

  const isFade = opts.effect === 'fadeIn' || opts.effect === 'fadeOut';
  const isEntrance = preset.presetClass === 'entr';

  const effectChildren: XmlElement[] = [];
  // Visibility kick: entrance reveals, exit hides.
  effectChildren.push(buildSetVisibility(spid, isEntrance));
  if (isFade) {
    effectChildren.push(buildOpacityAnim(spid, duration, isEntrance));
  }

  // cTn id=5 — the effect node.
  const effectCTn = elem(NAME_C_TN, {
    attrs: [
      attr(ATTR_ID, '5'),
      attr(ATTR_PRESET_ID, String(preset.presetId)),
      attr(ATTR_PRESET_CLASS, preset.presetClass),
      attr(ATTR_PRESET_SUBTYPE, String(preset.presetSubtype)),
      attr(ATTR_FILL, 'hold'),
      attr(ATTR_GRP_ID, '0'),
      attr(ATTR_NODE_TYPE, 'clickEffect'),
    ],
    children: [
      elem(NAME_ST_COND_LST, {
        children: [elem(NAME_COND, { attrs: [attr(ATTR_DELAY, '0')] })],
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

  // cTn id=3 — the indefinite wrapper (waiting for click).
  const indefiniteCTn = elem(NAME_C_TN, {
    attrs: [attr(ATTR_ID, '3'), attr(ATTR_FILL, 'hold')],
    children: [
      elem(NAME_ST_COND_LST, {
        children: [elem(NAME_COND, { attrs: [attr(ATTR_DELAY, 'indefinite')] })],
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

  // bldLst entry — required for PowerPoint to render the effect.
  const bldLst = elem(NAME_BLD_LST, {
    children: [
      elem(NAME_BLD_P, {
        attrs: [attr(ATTR_SPID, String(spid)), attr(ATTR_GRP_ID, '0')],
      }),
    ],
  });

  return elem(NAME_TIMING, { children: [tnLst, bldLst] });
};
