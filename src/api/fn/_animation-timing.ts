// Reading a slide's `<p:timing>` animation tree, and measuring its clock.
//
// Every animation read path goes through here, so the library cannot disagree
// with itself about what a timing tree contains. Callers get the effect nodes
// in document order together with the elements behind them, which is what the
// editing paths need. The write paths share the clock at the end of the file
// for the same reason: adding an effect, retiming one and reordering a click
// stop all have to agree about where a group starts.
//
// The parser is deliberately tolerant: a node it does not understand is
// reported as read-only rather than dropped, and one unknown node never hides
// the known effects around it. Anything the editing paths could corrupt —
// an effect spanning several targets, a duplicated `<p:cTn id>` — is reported
// with `editable: false` instead of a handle.

import {
  ANIMATION_DIRECTIONS,
  type AnimationDirection,
  type AnimationEffect,
  FULL_TURN,
  isMediaTimingNode,
  trailingHideDelayMs,
} from '../../internal/presentationml/index.ts';
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
const NAME_C_BHVR = qname('p', 'cBhvr', NS.pml);
const NAME_TO = qname('p', 'to', NS.pml);
const NAME_STR_VAL = qname('p', 'strVal', NS.pml);

const ATTR_ID = qname('', 'id', '');
const ATTR_DUR = qname('', 'dur', '');
const ATTR_DELAY = qname('', 'delay', '');
const ATTR_EVT = qname('', 'evt', '');
const ATTR_NODE_TYPE = qname('', 'nodeType', '');
const ATTR_PRESET_ID = qname('', 'presetID', '');
const ATTR_PRESET_CLASS = qname('', 'presetClass', '');
const ATTR_PRESET_SUBTYPE = qname('', 'presetSubtype', '');
const ATTR_GRP_ID = qname('', 'grpId', '');
const ATTR_SPID = qname('', 'spid', '');
const ATTR_BUILD = qname('', 'build', '');
const ATTR_BLD_LVL = qname('', 'bldLvl', '');
const ATTR_ST = qname('', 'st', '');
const ATTR_END = qname('', 'end', '');
const ATTR_FILL = qname('', 'fill', '');
const ATTR_BY = qname('', 'by', '');
const ATTR_FROM = qname('', 'from', '');
const ATTR_TO = qname('', 'to', '');
const ATTR_VAL = qname('', 'val', '');

/** `fill` values that leave a timing node's value in place after it has run. */
const HOLDING_FILLS = new Set(['hold', 'freeze']);

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
 * What becomes of the value an effect animates once it has run, from the
 * `fill` on its behaviours (ST_TLTimeNodeFillType: `remove | freeze | hold |
 * transition`). `'unstated'` is not `'held'`: the schema gives the attribute
 * no default, so a tree that omits it has not said.
 */
export type AnimationValueAfterEnd = 'held' | 'removed' | 'unstated';

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
  /** `null` for a preset outside the `AnimationEffect` tokens. */
  readonly effect: AnimationEffect | null;
  /**
   * Which edge of the slide a fly comes from or leaves by. `null` for every
   * other effect — the preset states no direction, so neither does this.
   */
  readonly direction: AnimationDirection | null;
  readonly presetId: number | null;
  readonly presetClass: string | null;
  /** `'unknown'` when the node type is not one this library models. */
  readonly start: AnimationStart;
  /** `null` when the tree does not state one — never defaulted. */
  readonly durationMs: number | null;
  /** `null` when the tree does not state one, or the delay is indefinite. */
  readonly delayMs: number | null;
  /**
   * What the tree says becomes of the value this effect animates once it has
   * run. It decides nothing on its own — `playable` already requires that the
   * behaviour saying whether the shape is on the slide keeps its value — but
   * it is what two effects animating one object at the same time turn on: the
   * one that ends first either stays above the other or gives way to it. A
   * player that finds `'unstated'` there must not choose for the file.
   */
  readonly valueAfterEnd: AnimationValueAfterEnd;
  /** `<p:bldP build="p">` — the text body is revealed paragraph by paragraph. */
  readonly buildByParagraph: boolean;
  readonly buildLevel: number | null;
  /**
   * Which `<p:seq>` the step lives in, named after its `nodeType`. Only
   * `'mainSeq'` advances on the slide's own clicks: an `'interactiveSeq'`
   * sequence fires when the viewer clicks the shape it is bound to, so a
   * player that folded one into the click order would run it at the wrong
   * time — and swallow a click the slide owes its main sequence.
   */
  readonly sequence: AnimationSequenceKind;
  /**
   * `true` when a player can run this step as part of the slide's click
   * order: it is in the main sequence, starts in a way we model, animates a
   * single target we model, and uses a preset we know how to render. A step
   * that is not playable must be shown as such, never folded into the normal
   * progression.
   */
  readonly playable: boolean;
  /**
   * `true` when this library can safely move, remove or retime the step:
   * `playable`, plus a unique handle to address it by. Editability is not
   * playability — a step can be one without the other as either side grows.
   */
  readonly editable: boolean;
}

/**
 * Which `<p:seq>` an effect belongs to, named after its `nodeType`. The two
 * tokens are ST_TLTimeNodeType values; `'other'` covers a sequence whose node
 * type is absent or something else again.
 */
export type AnimationSequenceKind = 'mainSeq' | 'interactiveSeq' | 'other';

/** A parsed step plus the elements it came from, for the editing paths. */
export interface AnimationStepNode {
  readonly step: SlideAnimationStep;
  readonly cTn: XmlElement;
  readonly sequence: XmlElement;
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

const sequenceKind = (seq: XmlElement): AnimationSequenceKind => {
  const cTn = firstChildElement(seq, NAME_C_TN);
  const nodeType = cTn === null ? null : getAttrValue(cTn, ATTR_NODE_TYPE);
  if (nodeType === 'mainSeq') return 'mainSeq';
  if (nodeType === 'interactiveSeq') return 'interactiveSeq';
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

// The `<p:set>` that flips visibility is scaffolding around every preset, not
// the effect's own length, so it is not what a duration change should touch.
const isVisibilityKick = (cBhvr: XmlElement): boolean => {
  const attrNameLst = firstChildElement(cBhvr, NAME_ATTR_NAME_LST);
  const attrName = attrNameLst === null ? null : firstChildElement(attrNameLst, NAME_ATTR_NAME);
  return (attrName?.children.find((c) => c.kind === 'text')?.data ?? '') === VISIBILITY_ATTR_NAME;
};

// The duration lives on the behaviour's own cTn, not on the effect node, so we
// read the first behaviour that is not the visibility kick.
const readDuration = (step: XmlElement): number | null => {
  const own = intAttr(step, ATTR_DUR);
  if (own !== null) return own;
  for (const cBhvr of behavioursOf(step)) {
    if (isVisibilityKick(cBhvr)) continue;
    const cTn = firstChildElement(cBhvr, NAME_C_TN);
    const dur = cTn === null ? null : intAttr(cTn, ATTR_DUR);
    if (dur !== null) return dur;
  }
  return null;
};

/**
 * Whether the shape is left where the effect put it once the effect has run.
 *
 * `fill` says what becomes of a timing node's value when its active duration
 * ends: ST_TLTimeNodeFillType is `remove | freeze | hold | transition`, and
 * CT_TLCommonTimeNodeData makes the attribute optional with no default, so a
 * node that carries none does not state that its value stands either.
 *
 * The effect node is asked first: its fill covers everything beneath it. Then
 * the behaviour that says where the effect leaves the shape. When the effect
 * flips `style.visibility` that is the `<p:set>`: on or off the slide is what
 * this library models, and it does not depend on the opacity underneath. An
 * effect that only fades has no such node, and then the opacity *is* the
 * answer: taken away at the end, it would put the shape back at the opacity it
 * started from, which is the opposite of what an exit states.
 *
 * What a fade leaves behind still matters where two effects animate one object
 * at the same time, but that is not a fact about either effect on its own: it
 * needs the moment each of them runs and the elements each of them reaches.
 * This reports it as `valueAfterEnd` and leaves the question to whoever plays
 * the slide.
 */
const holdsValue = (step: XmlElement): boolean => {
  if (!HOLDING_FILLS.has(getAttrValue(step, ATTR_FILL) ?? '')) return false;
  const behaviours = behavioursOf(step);
  const visibility = behaviours.filter(isVisibilityKick);
  for (const cBhvr of visibility.length > 0 ? visibility : behaviours) {
    const cTn = firstChildElement(cBhvr, NAME_C_TN);
    if (cTn === null || !HOLDING_FILLS.has(getAttrValue(cTn, ATTR_FILL) ?? '')) return false;
  }
  return true;
};

/**
 * What the tree says becomes of the value an effect animates once it has run.
 *
 * `fill` is asked of every behaviour that animates something — the `<p:set>`
 * that flips visibility is scaffolding, and `holdsValue` above asks it
 * separately. An absent `fill` is not read as `hold`: the schema gives it no
 * default, so the tree has not said.
 */
const readValueAfterEnd = (step: XmlElement): AnimationValueAfterEnd => {
  let stated = true;
  for (const cBhvr of behavioursOf(step)) {
    if (isVisibilityKick(cBhvr)) continue;
    const cTn = firstChildElement(cBhvr, NAME_C_TN);
    const fill = cTn === null ? null : getAttrValue(cTn, ATTR_FILL);
    if (fill === null) stated = false;
    else if (!HOLDING_FILLS.has(fill)) return 'removed';
  }
  return stated ? 'held' : 'unstated';
};

/** The bit PowerPoint's directional presets set for each edge of the slide. */
const FLY_SUBTYPES: Record<AnimationDirection, number> = { top: 1, right: 2, bottom: 4, left: 8 };

interface PresetEffect {
  readonly effect: AnimationEffect;
  readonly direction: AnimationDirection | null;
}

/**
 * Which effect a `(presetClass, presetID, presetSubtype)` triple names.
 *
 * All three are read, because for several presets the subtype is what the
 * effect actually does: entrance preset 23 grows from nothing at subtype 16 and
 * shrinks in from four times the size at subtype 32, and preset 2 flies from a
 * different edge for each of its bits. Reading only the first two would report
 * one as the other and animate the shape the wrong way.
 *
 * `'-'` stands for a subtype the tree does not state. It is accepted only where
 * PowerPoint's gallery offers the preset no options at all, so there is a
 * single thing it could have meant; a fly or a zoom without a subtype is an
 * effect we cannot name, and neither are the diagonal flies (two bits at once)
 * this library does not write.
 */
const PRESET_EFFECTS: ReadonlyMap<string, PresetEffect> = (() => {
  const out = new Map<string, PresetEffect>();
  const plain = (cls: string, id: number, effect: AnimationEffect): void => {
    out.set(`${cls}:${id}:0`, { effect, direction: null });
    out.set(`${cls}:${id}:-`, { effect, direction: null });
  };
  plain('entr', 1, 'appear');
  plain('entr', 10, 'fadeIn');
  plain('exit', 1, 'disappear');
  plain('exit', 10, 'fadeOut');
  plain('emph', 8, 'spin');
  out.set('entr:23:16', { effect: 'zoomIn', direction: null });
  out.set('exit:23:32', { effect: 'zoomOut', direction: null });
  for (const direction of ANIMATION_DIRECTIONS) {
    const subtype = FLY_SUBTYPES[direction];
    out.set(`entr:2:${subtype}`, { effect: 'flyIn', direction });
    out.set(`exit:2:${subtype}`, { effect: 'flyOut', direction });
  }
  return out;
})();

/**
 * Whether a `spin` really turns the shape once, clockwise, and does nothing
 * else.
 *
 * The preset triple is not enough. `presetClass="emph" presetID="8"` is
 * PowerPoint's Spin whatever angle it turns through: the amount lives on
 * `<p:animRot>`, whose `by` / `from` / `to` are `a:ST_Angle` — sixtieth-
 * thousandths of a degree, and negative for a counter-clockwise turn. A tree
 * that says half a turn back the other way carries the same preset numbers as
 * one that says a full turn forwards, so reading the numbers alone would report
 * the second when the file says the first — and let an edit rewrite it as one.
 *
 * Only the shape this library writes is accepted: a single rotation behaviour,
 * `by` a full positive turn, and no `from` / `to` fixing where it starts or
 * ends. Anything else is an effect we read rather than name.
 */
const isFullClockwiseTurn = (step: XmlElement): boolean => {
  const rotations: XmlElement[] = [];
  const walk = (el: XmlElement): void => {
    if (isPml(el, 'cBhvr')) return;
    if (isPml(el, 'animRot')) {
      rotations.push(el);
      return;
    }
    for (const child of el.children) if (child.kind === 'element') walk(child);
  };
  for (const child of step.children) if (child.kind === 'element') walk(child);
  const rotation = rotations[0];
  if (rotations.length !== 1 || rotation === undefined) return false;
  if (behavioursOf(step).length !== 1) return false;
  if (getAttrValue(rotation, ATTR_FROM) !== null || getAttrValue(rotation, ATTR_TO) !== null) {
    return false;
  }
  return strictInt(getAttrValue(rotation, ATTR_BY)) === FULL_TURN;
};

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
  sequence: AnimationSequenceKind,
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
  const rawSubtype = getAttrValue(node, ATTR_PRESET_SUBTYPE);
  // A subtype that is stated but is not a plain integer is not "unstated": it
  // says something we cannot read, so it matches no entry rather than the one
  // for a preset written without a subtype at all.
  const subtype = rawSubtype === null ? '-' : (strictInt(rawSubtype) ?? 'unreadable');
  const named = PRESET_EFFECTS.get(`${presetClass}:${presetId}:${subtype}`) ?? null;
  // A preset whose numbers do not settle what it does is checked against the
  // behaviour that does.
  const preset = named?.effect === 'spin' && !isFullClockwiseTurn(node) ? null : named;
  const effect = preset?.effect ?? null;
  const playable =
    sequence === 'mainSeq' &&
    start !== 'unknown' &&
    target.kind !== 'unsupported' &&
    effect !== null &&
    holdsValue(node);

  return {
    id,
    target,
    targetShapeIds,
    effect,
    direction: preset?.direction ?? null,
    presetId,
    presetClass,
    start,
    durationMs: readDuration(node),
    delayMs: readStartDelay(node),
    valueAfterEnd: readValueAfterEnd(node),
    buildByParagraph: bldP !== undefined && getAttrValue(bldP, ATTR_BUILD) === 'p',
    buildLevel: bldP === undefined ? null : intAttr(bldP, ATTR_BLD_LVL),
    sequence,
    // A preset we cannot name may animate anything at all, and a step outside
    // the main sequence is triggered by something other than the slide's own
    // clicks, so neither may join the normal progression. Nor may one whose
    // `fill` does not say its value stands once it has run: an entrance that
    // is taken away again afterwards leaves the shape somewhere we do not
    // model, and playing it as a plain entrance would show what the deck hides.
    playable,
    // A preset we cannot name may also hang extra timing references off
    // behaviours we have never seen, so rewriting it is not demonstrably safe.
    editable: playable && id !== null,
  };
};

/**
 * Every animation effect on the slide, in document order, with the elements
 * behind it. Media time nodes are skipped — a clip's play controls live in the
 * same `<p:timing>` but are not animations.
 */
export const readSlideTiming = (slide: SlideData): AnimationStepNode[] => {
  const timing = findSlideTimingElement(slide);
  return timing === null ? [] : readTimingSteps(timing);
};

/**
 * The same read against a `<p:timing>` element directly, for the editing paths:
 * they assemble a change in a copy of the tree, which is not on a slide yet.
 */
export const readTimingSteps = (timing: XmlElement): AnimationStepNode[] => {
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
      });
    }
  }

  return out;
};

// ---------------------------------------------------------------------------
// The click stop's clock.
//
// Sibling `<p:par>` nodes are parallel — they all begin when their parent
// does — so `nodeType="afterEffect"` is a label for PowerPoint's UI, not a
// dependency a player could honour. What actually makes one group run after
// another is the start offset on the later group. Writing an effect, changing
// its duration or delay, and reordering a stop all have to agree about where
// that offset falls, so they all measure through here.
//
// Measuring is deliberately narrow: an authored slide can time an effect in
// ways this library does not model, and a guessed end would place the next
// effect at a moment PowerPoint never plays it. Anything outside the modelled
// shapes reports `null`, and the caller refuses rather than guessing.

// Whole milliseconds only. `indefinite`, a missing value or anything that is
// not a plain non-negative integer means we cannot place a node on the
// timeline.
const wholeMs = (raw: string | null): number | null => {
  if (raw === null || !/^\d+$/.test(raw)) return null;
  const n = Number(raw);
  return Number.isSafeInteger(n) ? n : null;
};

/**
 * The single `<p:cond>` that fixes when a node starts, or `null` when the node
 * does not start at a plain offset at all.
 *
 * Several conditions mean several triggers, and `evt` or a `<p:tn>` / `<p:rtn>`
 * / `<p:tgtEl>` child ties the start to another node's lifetime — neither is a
 * number of milliseconds, and reading the `delay` attribute beside one as if it
 * were the offset would place the node at a time PowerPoint never plays it.
 */
export const offsetCond = (cTn: XmlElement): XmlElement | null => {
  const stCondLst = firstChildElement(cTn, NAME_ST_COND_LST);
  if (stCondLst === null) return null;
  const conds = stCondLst.children.filter(
    (c): c is XmlElement => c.kind === 'element' && isPml(c, 'cond'),
  );
  if (conds.length !== 1) return null;
  const cond = conds[0]!;
  if (getAttrValue(cond, ATTR_EVT) !== null) return null;
  if (cond.children.some((c) => c.kind === 'element')) return null;
  return cond;
};

/** How long after its parent a node starts. A node with no condition starts with it. */
const startOffsetMs = (cTn: XmlElement): number | null => {
  if (firstChildElement(cTn, NAME_ST_COND_LST) === null) return 0;
  const cond = offsetCond(cTn);
  return cond === null ? null : wholeMs(getAttrValue(cond, ATTR_DELAY));
};

/**
 * Attributes that repeat or rescale a time node (CT_TLCommonTimeNodeData).
 * We measure a node by its `dur`, so any of these means the number we computed
 * is not how long PowerPoint runs it. `accel` / `decel` are shares of `dur`
 * and leave the total alone, so they are not here.
 */
const RESCALING_ATTRS = new Set(['repeatCount', 'repeatDur', 'spd', 'autoRev']);

const isRescaled = (cTn: XmlElement): boolean =>
  cTn.attrs.some((a) => a.name.namespaceURI === '' && RESCALING_ATTRS.has(a.name.localName));

/**
 * Time nodes nested inside an effect. Their children start when *they* do
 * rather than when the effect does, and `<p:iterate>` staggers them further
 * per letter or paragraph, so measuring the behaviours underneath as if they
 * were the effect's own would misplace every one of them.
 */
const NESTED_TIMELINES = new Set(['par', 'seq', 'excl', 'iterate']);

/**
 * How long after its own start an effect is still running, or `null` when the
 * tree does not say. The 1ms visibility kick counts — for `appear` it is the
 * whole effect.
 *
 * Every behaviour has to be measurable. One that runs indefinitely, states no
 * length, repeats, or hangs off a nested timeline decides the answer on its
 * own: falling back to the readable siblings' maximum would report an effect
 * as shorter than it plays, and everything placed after it would start early.
 */
const effectSpanMs = (effectPar: XmlElement): number | null => {
  const effectCTn = firstChildElement(effectPar, NAME_C_TN);
  if (effectCTn === null || isRescaled(effectCTn)) return null;
  const start = startOffsetMs(effectCTn);
  if (start === null) return null;

  let longest: number | null = null;
  const walk = (el: XmlElement): boolean => {
    if (isPml(el, 'cBhvr')) {
      const cTn = firstChildElement(el, NAME_C_TN);
      if (cTn === null || isRescaled(cTn)) return false;
      const dur = wholeMs(getAttrValue(cTn, ATTR_DUR));
      const delay = startOffsetMs(cTn);
      if (dur === null || delay === null) return false;
      longest = Math.max(longest ?? 0, delay + dur);
      return true;
    }
    for (const child of el.children) {
      if (child.kind !== 'element') continue;
      if (child.name.namespaceURI === NS.pml && NESTED_TIMELINES.has(child.name.localName)) {
        return false;
      }
      if (!walk(child)) return false;
    }
    return true;
  };
  if (!walk(effectCTn) || longest === null) return null;
  return start + longest;
};

/**
 * When every effect in `groupPar` has finished, measured from the start of the
 * click stop that holds it. `null` when any of them runs for a length this
 * library cannot measure, which is what makes "after this one" unanswerable.
 */
export const groupEndMs = (groupPar: XmlElement): number | null => {
  const groupCTn = firstChildElement(groupPar, NAME_C_TN);
  if (groupCTn === null || isRescaled(groupCTn)) return null;
  const start = startOffsetMs(groupCTn);
  if (start === null) return null;

  // Anything in the group that is not a plain `<p:par>` effect is a structure
  // we have not measured, so the group's end is not ours to state.
  const childTnLst = firstChildElement(groupCTn, NAME_CHILD_TN_LST);
  if (childTnLst === null) return null;
  const effectPars = childTnLst.children.filter((c): c is XmlElement => c.kind === 'element');
  if (!effectPars.every((c) => isPml(c, 'par'))) return null;

  let longest: number | null = null;
  for (const effectPar of effectPars) {
    const span = effectSpanMs(effectPar);
    if (span === null) return null;
    longest = Math.max(longest ?? 0, span);
  }
  return longest === null ? null : start + longest;
};

/**
 * Moves `groupPar` to start `delayMs` after its click stop. Returns false when
 * the group does not start at a plain offset, so the caller leaves it alone
 * rather than overwriting a condition it does not understand.
 */
export const setGroupStartOffset = (groupPar: XmlElement, delayMs: number): boolean => {
  const cTn = firstChildElement(groupPar, NAME_C_TN);
  const cond = cTn === null ? null : offsetCond(cTn);
  if (cond === null) return false;
  cond.attrs = cond.attrs.map((a) =>
    a.name.namespaceURI === '' && a.name.localName === 'delay'
      ? { ...a, value: String(delayMs) }
      : a,
  );
  return true;
};

const setTimeAttr = (el: XmlElement, name: string, value: string): void => {
  const has = el.attrs.some((a) => a.name.namespaceURI === '' && a.name.localName === name);
  el.attrs = has
    ? el.attrs.map((a) =>
        a.name.namespaceURI === '' && a.name.localName === name ? { ...a, value } : a,
      )
    : [...el.attrs, { name: qname('', name, ''), value }];
};

/**
 * Sets how long after its group an effect waits, in place. `false` when the
 * effect does not start at a plain offset: the condition there ties it to
 * something else, and overwriting the delay beside it would not mean what the
 * caller asked for.
 */
export const setEffectDelayMs = (effectCTn: XmlElement, delayMs: number): boolean => {
  const cond = offsetCond(effectCTn);
  if (cond === null) return false;
  setTimeAttr(cond, 'delay', String(delayMs));
  return true;
};

/**
 * Sets how long an effect runs, in place, leaving everything else about it
 * alone.
 *
 * An effect may animate through several behaviours at once — a fly drives both
 * axes, a zoom both sides — and they run together, so one length covers all of
 * them. `false` when they do not already agree on one: an effect whose
 * behaviours run for different lengths has no single length to set, and picking
 * one would silently retime the rest of it.
 *
 * An effect with no timed behaviour — `appear` and `disappear` write only the
 * visibility kick — has no duration to change and is left as it is.
 */
export const setEffectDurationMs = (effectCTn: XmlElement, durationMs: number): boolean => {
  const behaviours = behavioursOf(effectCTn);
  const timed = behaviours.filter((b) => !isVisibilityKick(b));
  const cTns = timed.map((b) => firstChildElement(b, NAME_C_TN));
  if (cTns.some((cTn) => cTn === null)) return false;
  const durations = new Set(cTns.map((cTn) => getAttrValue(cTn!, ATTR_DUR)));
  if (durations.size > 1) return false;
  // Nothing timed to set — `appear` and `disappear` write only the visibility
  // kick, and it is the whole of the effect rather than something trailing it.
  if (cTns.length === 0) return true;
  const was = wholeMs([...durations][0] ?? null);
  for (const cTn of cTns) setTimeAttr(cTn!, 'dur', String(durationMs));
  // An exit stays on the slide until its motion is over, so the `<p:set>` that
  // hides the shape trails that motion by the millisecond it takes itself, and
  // has to move when the end does.
  //
  // It is found by where it stands rather than by standing anywhere but zero:
  // an exit written at a duration of one millisecond or less already hides at
  // zero, so "not at zero" would strand it there on the next change. A kick
  // somewhere else entirely belongs to timing this library did not write, and
  // moving it would say something about the effect the file never did.
  const trailing = was === null ? null : String(trailingHideDelayMs(was));
  if (trailing === null) return true;
  for (const hide of hideSets(effectCTn)) {
    const cTn = firstChildElement(hide, NAME_C_TN);
    const cond = cTn === null ? null : offsetCond(cTn);
    if (cond === null || getAttrValue(cond, ATTR_DELAY) !== trailing) continue;
    setTimeAttr(cond, 'delay', String(trailingHideDelayMs(durationMs)));
  }
  return true;
};

/**
 * The `<p:cBhvr>` of every `<p:set>` in the effect that takes its target off the
 * slide. A `<p:set>` putting it *on* is an entrance's opening kick and never
 * trails anything.
 */
const hideSets = (step: XmlElement): XmlElement[] => {
  const out: XmlElement[] = [];
  const walk = (el: XmlElement): void => {
    if (isPml(el, 'cBhvr')) return;
    if (isPml(el, 'set')) {
      const cBhvr = firstChildElement(el, NAME_C_BHVR);
      const to = firstChildElement(el, NAME_TO);
      const strVal = to === null ? null : firstChildElement(to, NAME_STR_VAL);
      if (
        cBhvr !== null &&
        isVisibilityKick(cBhvr) &&
        strVal !== null &&
        getAttrValue(strVal, ATTR_VAL) === 'hidden'
      ) {
        out.push(cBhvr);
      }
      return;
    }
    for (const child of el.children) if (child.kind === 'element') walk(child);
  };
  for (const child of step.children) if (child.kind === 'element') walk(child);
  return out;
};
