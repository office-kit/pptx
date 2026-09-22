// Reading a slide's `<p:timing>` animation tree.
//
// Every animation read path goes through here, so the library cannot disagree
// with itself about what a timing tree contains. Callers get the effect nodes
// in document order together with the elements behind them, which is what the
// editing paths need.
//
// The parser is deliberately tolerant: a node it does not understand is
// reported as read-only rather than dropped, and one unknown node never hides
// the known effects around it. Anything the editing paths could corrupt —
// an effect spanning several targets, a duplicated `<p:cTn id>` — is reported
// with `editable: false` instead of a handle.

import { type AnimationEffect, isMediaTimingNode } from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';
import { SLIDE_DOCUMENT, type SlideData } from '../_internal-symbols.ts';
import { rootChildTnLst } from './_media-timing.ts';

const NAME_TIMING = qname('p', 'timing', NS.pml);
const NAME_BLD_LST = qname('p', 'bldLst', NS.pml);
const NAME_ST_COND_LST = qname('p', 'stCondLst', NS.pml);
const NAME_COND = qname('p', 'cond', NS.pml);
const NAME_C_TN = qname('p', 'cTn', NS.pml);
const NAME_CHILD_TN_LST = qname('p', 'childTnLst', NS.pml);
const NAME_ATTR_NAME_LST = qname('p', 'attrNameLst', NS.pml);
const NAME_ATTR_NAME = qname('p', 'attrName', NS.pml);
const NAME_TGT_EL = qname('p', 'tgtEl', NS.pml);
const NAME_SP_TGT = qname('p', 'spTgt', NS.pml);
const NAME_TX_EL = qname('p', 'txEl', NS.pml);
const NAME_P_RG = qname('p', 'pRg', NS.pml);

const ATTR_ID = qname('', 'id', '');
const ATTR_DUR = qname('', 'dur', '');
const ATTR_DELAY = qname('', 'delay', '');
const ATTR_NODE_TYPE = qname('', 'nodeType', '');
const ATTR_PRESET_ID = qname('', 'presetID', '');
const ATTR_PRESET_CLASS = qname('', 'presetClass', '');
const ATTR_GRP_ID = qname('', 'grpId', '');
const ATTR_SPID = qname('', 'spid', '');
const ATTR_BUILD = qname('', 'build', '');
const ATTR_BLD_LVL = qname('', 'bldLvl', '');
const ATTR_ST = qname('', 'st', '');
const ATTR_END = qname('', 'end', '');

// The `<p:set>` that flips `style.visibility` is scaffolding around every
// preset, not the effect's own timing, so it never supplies the duration.
const VISIBILITY_ATTR_NAME = 'style.visibility';

// CT_TLTimeNodeParallel / Sequence children that animate something, as opposed
// to the `<p:par>` wrappers that only group. A node holding one of these is an
// effect even when it carries no preset attributes at all.
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

const STEP_NODE_TYPES = new Set(['clickEffect', 'withEffect', 'afterEffect']);

/** When a step runs relative to the one before it. */
export type AnimationStart = 'click' | 'withPrevious' | 'afterPrevious' | 'unknown';

/**
 * What an effect animates. `unsupported` covers targets this library does not
 * model — charts, OLE parts, sub-shapes, and composite effects whose
 * behaviours do not all animate the same thing. The step is still listed so a
 * caller can show it, but it cannot be edited.
 */
export type AnimationTarget =
  | { readonly kind: 'shape'; readonly shapeId: number }
  | {
      readonly kind: 'paragraphs';
      readonly shapeId: number;
      readonly firstParagraph: number;
      readonly lastParagraph: number;
    }
  | { readonly kind: 'unsupported'; readonly shapeId: number | null };

/** One animation effect on a slide, as read from `<p:timing>`. */
export interface SlideAnimationStep {
  /**
   * How the editing APIs address this step: the effect's `<p:cTn id>`.
   * `null` when the node carries no id, the id is not a plain integer, or the
   * same id appears more than once in the tree — none of which is a safe
   * handle, so such a step is read-only.
   */
  readonly id: number | null;
  readonly target: AnimationTarget;
  /**
   * Every shape any of the step's behaviours animates, in document order.
   * A composite effect that drives two shapes reports both here even though
   * `target` refuses to name one of them as *the* target.
   */
  readonly targetShapeIds: readonly number[];
  /** `null` for presets outside the four `AnimationEffect` tokens. */
  readonly effect: AnimationEffect | null;
  readonly presetId: number | null;
  readonly presetClass: string | null;
  /** `'unknown'` when the node type is not one this library models. */
  readonly start: AnimationStart;
  /** `null` when the tree does not state one — never defaulted. */
  readonly durationMs: number | null;
  /** `null` when the tree does not state one, or the delay is indefinite. */
  readonly delayMs: number | null;
  /** `<p:bldP build="p">` — the text body is revealed paragraph by paragraph. */
  readonly buildByParagraph: boolean;
  readonly buildLevel: number | null;
  /**
   * `true` when this library can safely move, remove or retime the step: it
   * has a unique handle, sits in the main sequence, starts in a way we model,
   * animates a single target we model, and uses a preset we recognise. A step
   * built from a preset we cannot name is still listed, but read-only: we do
   * not know what else its behaviours reference, so we must not rewrite it.
   */
  readonly editable: boolean;
}

/** A parsed step plus the elements it came from, for the editing paths. */
export interface AnimationStepNode {
  readonly step: SlideAnimationStep;
  readonly cTn: XmlElement;
  readonly sequence: XmlElement;
  readonly sequenceKind: 'mainSeq' | 'interactive' | 'other';
}

const isPml = (el: XmlElement, local: string): boolean =>
  el.name.namespaceURI === NS.pml && el.name.localName === local;

// `Number.parseInt` accepts trailing garbage ("4abc" -> 4). An attribute that
// is not a plain integer is not a value we may act on, so reject it outright.
const strictInt = (raw: string | null): number | null => {
  if (raw === null || !/^[+-]?\d+$/.test(raw)) return null;
  const n = Number(raw);
  return Number.isSafeInteger(n) ? n : null;
};

const intAttr = (el: XmlElement, name: ReturnType<typeof qname>): number | null =>
  strictInt(getAttrValue(el, name));

export const findSlideTimingElement = (slide: SlideData): XmlElement | null =>
  firstChildElement(slide[SLIDE_DOCUMENT].root, NAME_TIMING);

// `<p:bldP spid grpId build bldLvl>` keyed by the pair that ties it to an
// effect. PowerPoint groups a shape's build with its effect through grpId.
const readBuilds = (timing: XmlElement): Map<string, XmlElement> => {
  const out = new Map<string, XmlElement>();
  const bldLst = firstChildElement(timing, NAME_BLD_LST);
  if (bldLst === null) return out;
  for (const child of bldLst.children) {
    if (child.kind !== 'element') continue;
    const spid = getAttrValue(child, ATTR_SPID);
    const grpId = getAttrValue(child, ATTR_GRP_ID);
    if (spid !== null) out.set(`${spid}:${grpId ?? '0'}`, child);
  }
  return out;
};

const sequenceKind = (seq: XmlElement): AnimationStepNode['sequenceKind'] => {
  const cTn = firstChildElement(seq, NAME_C_TN);
  const nodeType = cTn === null ? null : getAttrValue(cTn, ATTR_NODE_TYPE);
  if (nodeType === 'mainSeq') return 'mainSeq';
  if (nodeType === 'interactive') return 'interactive';
  return 'other';
};

const holdsBehaviour = (cTn: XmlElement): boolean => {
  const childTnLst = firstChildElement(cTn, NAME_CHILD_TN_LST);
  if (childTnLst === null) return false;
  return childTnLst.children.some(
    (c) =>
      c.kind === 'element' &&
      c.name.namespaceURI === NS.pml &&
      BEHAVIOUR_LOCALS.has(c.name.localName),
  );
};

// A step is a time node that stands for one effect, whether or not PowerPoint
// wrote preset attributes on it. Behaviour time nodes (inside `<p:cBhvr>`)
// describe *how* an effect animates, so they are never steps — the walk skips
// those subtrees entirely.
const isStepNode = (cTn: XmlElement): boolean =>
  getAttrValue(cTn, ATTR_PRESET_ID) !== null ||
  getAttrValue(cTn, ATTR_PRESET_CLASS) !== null ||
  STEP_NODE_TYPES.has(getAttrValue(cTn, ATTR_NODE_TYPE) ?? '') ||
  holdsBehaviour(cTn);

const collectStepNodes = (el: XmlElement, into: XmlElement[]): void => {
  if (isPml(el, 'cBhvr')) return;
  if (isPml(el, 'cTn') && isStepNode(el)) {
    into.push(el);
    return;
  }
  for (const child of el.children) if (child.kind === 'element') collectStepNodes(child, into);
};

const behavioursOf = (step: XmlElement): XmlElement[] => {
  const out: XmlElement[] = [];
  const walk = (el: XmlElement): void => {
    if (isPml(el, 'cBhvr')) {
      out.push(el);
      return;
    }
    for (const child of el.children) if (child.kind === 'element') walk(child);
  };
  for (const child of step.children) if (child.kind === 'element') walk(child);
  return out;
};

const targetOfSpTgt = (spTgt: XmlElement): AnimationTarget => {
  const shapeId = intAttr(spTgt, ATTR_SPID);
  if (shapeId === null) return { kind: 'unsupported', shapeId: null };

  const elements = spTgt.children.filter((c): c is XmlElement => c.kind === 'element');
  if (elements.length === 0) return { kind: 'shape', shapeId };

  const txEl = firstChildElement(spTgt, NAME_TX_EL);
  const pRg = txEl === null ? null : firstChildElement(txEl, NAME_P_RG);
  if (elements.length === 1 && pRg !== null) {
    const first = intAttr(pRg, ATTR_ST);
    const last = intAttr(pRg, ATTR_END);
    if (first !== null && last !== null && first <= last) {
      return { kind: 'paragraphs', shapeId, firstParagraph: first, lastParagraph: last };
    }
  }
  // `<p:graphicEl>`, `<p:oleChartEl>`, `<p:subSp>` and a `<p:txEl>` that
  // selects by character range rather than paragraph all land here.
  return { kind: 'unsupported', shapeId };
};

const targetKey = (target: AnimationTarget): string =>
  target.kind === 'paragraphs'
    ? `p:${target.shapeId}:${target.firstParagraph}:${target.lastParagraph}`
    : `${target.kind}:${target.shapeId ?? ''}`;

// Every behaviour in an effect must animate the same thing for the effect to
// be one this library may move or delete: a composite effect that also drives
// a second shape would lose that shape if we treated it as belonging to the
// first one.
const readTargets = (
  step: XmlElement,
): { target: AnimationTarget; targetShapeIds: readonly number[] } => {
  const targets = behavioursOf(step).flatMap((cBhvr) => {
    const tgtEl = firstChildElement(cBhvr, NAME_TGT_EL);
    if (tgtEl === null) return [];
    const spTgt = firstChildElement(tgtEl, NAME_SP_TGT);
    return [
      spTgt === null ? ({ kind: 'unsupported', shapeId: null } as const) : targetOfSpTgt(spTgt),
    ];
  });
  const targetShapeIds = [
    ...new Set(targets.flatMap((t) => (t.shapeId === null ? [] : [t.shapeId]))),
  ];
  if (targets.length === 0)
    return { target: { kind: 'unsupported', shapeId: null }, targetShapeIds };
  const first = targets[0]!;
  const key = targetKey(first);
  if (targets.some((t) => targetKey(t) !== key)) {
    return { target: { kind: 'unsupported', shapeId: null }, targetShapeIds };
  }
  return { target: first, targetShapeIds };
};

const readStartDelay = (step: XmlElement): number | null => {
  const stCondLst = firstChildElement(step, NAME_ST_COND_LST);
  const cond = stCondLst === null ? null : firstChildElement(stCondLst, NAME_COND);
  return cond === null ? null : intAttr(cond, ATTR_DELAY);
};

// The duration lives on the behaviour's own cTn, not on the effect node, so we
// read the first behaviour that is not the visibility kick.
const readDuration = (step: XmlElement): number | null => {
  const own = intAttr(step, ATTR_DUR);
  if (own !== null) return own;
  for (const cBhvr of behavioursOf(step)) {
    const attrNameLst = firstChildElement(cBhvr, NAME_ATTR_NAME_LST);
    const attrName = attrNameLst === null ? null : firstChildElement(attrNameLst, NAME_ATTR_NAME);
    const text = attrName?.children.find((c) => c.kind === 'text')?.data ?? '';
    if (text === VISIBILITY_ATTR_NAME) continue;
    const cTn = firstChildElement(cBhvr, NAME_C_TN);
    const dur = cTn === null ? null : intAttr(cTn, ATTR_DUR);
    if (dur !== null) return dur;
  }
  return null;
};

const PRESET_EFFECTS: ReadonlyArray<readonly [string, number, AnimationEffect]> = [
  ['entr', 1, 'appear'],
  ['entr', 10, 'fadeIn'],
  ['exit', 1, 'disappear'],
  ['exit', 10, 'fadeOut'],
];

const readStart = (step: XmlElement): AnimationStart => {
  switch (getAttrValue(step, ATTR_NODE_TYPE)) {
    case 'clickEffect':
      return 'click';
    case 'withEffect':
      return 'withPrevious';
    case 'afterEffect':
      return 'afterPrevious';
    default:
      return 'unknown';
  }
};

// Ids that occur more than once cannot address a single node, so they are not
// handles. Counted across the whole `<p:timing>`, behaviour nodes included.
const countCTnIds = (timing: XmlElement): Map<number, number> => {
  const counts = new Map<number, number>();
  const walk = (el: XmlElement): void => {
    if (isPml(el, 'cTn')) {
      const id = intAttr(el, ATTR_ID);
      if (id !== null) counts.set(id, (counts.get(id) ?? 0) + 1);
    }
    for (const child of el.children) if (child.kind === 'element') walk(child);
  };
  walk(timing);
  return counts;
};

const toStep = (
  node: XmlElement,
  builds: Map<string, XmlElement>,
  sequence: AnimationStepNode['sequenceKind'],
  idCounts: Map<number, number>,
): SlideAnimationStep => {
  const presetClass = getAttrValue(node, ATTR_PRESET_CLASS);
  const presetId = intAttr(node, ATTR_PRESET_ID);
  const { target, targetShapeIds } = readTargets(node);
  const start = readStart(node);
  const rawId = intAttr(node, ATTR_ID);
  const id = rawId !== null && idCounts.get(rawId) === 1 ? rawId : null;

  const grpId = getAttrValue(node, ATTR_GRP_ID) ?? '0';
  const bldP = target.kind === 'unsupported' ? undefined : builds.get(`${target.shapeId}:${grpId}`);
  const effect =
    PRESET_EFFECTS.find(([cls, pid]) => cls === presetClass && pid === presetId)?.[2] ?? null;

  return {
    id,
    target,
    targetShapeIds,
    effect,
    presetId,
    presetClass,
    start,
    durationMs: readDuration(node),
    delayMs: readStartDelay(node),
    buildByParagraph: bldP !== undefined && getAttrValue(bldP, ATTR_BUILD) === 'p',
    buildLevel: bldP === undefined ? null : intAttr(bldP, ATTR_BLD_LVL),
    // A preset we cannot name may hang extra timing references off behaviours
    // we have never seen, so retiming or moving it is not demonstrably safe.
    // It stays read-only; the recognised steps around it do not.
    editable:
      id !== null &&
      sequence === 'mainSeq' &&
      start !== 'unknown' &&
      target.kind !== 'unsupported' &&
      effect !== null,
  };
};

/**
 * Every animation effect on the slide, in document order, with the elements
 * behind it. Media time nodes are skipped — a clip's play controls live in the
 * same `<p:timing>` but are not animations.
 */
export const readSlideTiming = (slide: SlideData): AnimationStepNode[] => {
  const timing = findSlideTimingElement(slide);
  if (timing === null) return [];
  const rootList = rootChildTnLst(timing);
  if (rootList === null) return [];
  const builds = readBuilds(timing);
  const idCounts = countCTnIds(timing);

  const out: AnimationStepNode[] = [];
  for (const child of rootList.children) {
    if (child.kind !== 'element' || isMediaTimingNode(child)) continue;
    const kind = sequenceKind(child);
    const nodes: XmlElement[] = [];
    collectStepNodes(child, nodes);
    for (const node of nodes) {
      out.push({
        step: toStep(node, builds, kind, idCounts),
        cTn: node,
        sequence: child,
        sequenceKind: kind,
      });
    }
  }
  return out;
};
