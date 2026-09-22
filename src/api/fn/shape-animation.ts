// Slide animations.

import {
  type AnimationEffect,
  type AnimationOptions,
  type AnimationStartCondition,
  buildSingleEffectTiming,
  buildTimingRoot,
} from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
  allChildElements,
  elem,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';
import {
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  SLIDE_DOCUMENT,
  SLIDE_SHAPES,
  type SlideData,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import {
  type AnimationSequenceKind,
  type AnimationStart,
  type AnimationTarget,
  type SlideAnimationStep,
  findSlideTimingElement,
  readSlideTiming,
} from './_animation-timing.ts';
import { commitSlideData, refreshSlideData } from './_helpers.ts';
import { maxCTnId, mediaTimingNodes, rootChildTnLst } from './_media-timing.ts';
// ---------------------------------------------------------------------------
// Animations (one effect per call).
//
// Current scope: each call adds one entrance or exit effect and merges it into
// whatever timing tree the slide already has, so a slide can carry several
// effects across several shapes and click stops. Emphasis and motion presets,
// per-paragraph builds, and editing or reordering an effect that is already
// there are not modelled yet.

export type { AnimationEffect, AnimationOptions, AnimationStartCondition };
export type { AnimationSequenceKind, AnimationStart, AnimationTarget, SlideAnimationStep };

const ATTR_ID_FN = qname('', 'id', '');

const removeExistingTiming = (slide: SlideData): void => {
  slide[SLIDE_DOCUMENT].root.children = slide[SLIDE_DOCUMENT].root.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'timing'),
  );
};

const findTiming = (slide: SlideData): XmlElement | null =>
  slide[SLIDE_DOCUMENT].root.children.find(
    (c): c is XmlElement =>
      c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'timing',
  ) ?? null;

const insertTimingAtEnd = (slide: SlideData, timing: XmlElement): void => {
  // Schema ordering: `<p:timing>` is one of the last children of `<p:sld>`
  // (after cSld, clrMapOvr, transition). Appending to the end of
  // `<p:sld>` keeps the file valid.
  slide[SLIDE_DOCUMENT].root.children.push(timing);
};

// Depth-first search for the first descendant (or self) matching `pred`.
const findDescendant = (el: XmlElement, pred: (e: XmlElement) => boolean): XmlElement | null => {
  if (pred(el)) return el;
  for (const c of el.children) {
    if (c.kind === 'element') {
      const found = findDescendant(c, pred);
      if (found) return found;
    }
  }
  return null;
};

const isPml = (el: XmlElement, local: string): boolean =>
  el.name.namespaceURI === NS.pml && el.name.localName === local;

const shiftCTnIds = (el: XmlElement, offset: number): void => {
  const walk = (e: XmlElement): void => {
    if (isPml(e, 'cTn')) {
      const raw = getAttrValue(e, ATTR_ID_FN);
      const n = raw === null ? Number.NaN : Number.parseInt(raw, 10);
      if (Number.isFinite(n)) {
        e.attrs = e.attrs.map((a) =>
          a.name.namespaceURI === '' && a.name.localName === 'id'
            ? { ...a, value: String(n + offset) }
            : a,
        );
      }
    }
    for (const c of e.children) if (c.kind === 'element') walk(c);
  };
  walk(el);
};

// Largest numeric `grpId` attribute anywhere in the tree (-1 when none), so a
// freshly-merged build group can take max+1 and never collide with an existing
// one even when the authored grpIds have gaps.
const maxGrpId = (el: XmlElement): number => {
  let max = -1;
  const walk = (e: XmlElement): void => {
    const raw = getAttrValue(e, qname('', 'grpId', ''));
    if (raw !== null) {
      const n = Number.parseInt(raw, 10);
      if (Number.isFinite(n) && n > max) max = n;
    }
    for (const c of e.children) if (c.kind === 'element') walk(c);
  };
  walk(el);
  return max;
};

const setGrpId = (el: XmlElement, grpId: string): void => {
  el.attrs = el.attrs.map((a) =>
    a.name.namespaceURI === '' && a.name.localName === 'grpId' ? { ...a, value: grpId } : a,
  );
};

const isMainSeqCTn = (e: XmlElement): boolean =>
  isPml(e, 'cTn') && getAttrValue(e, qname('', 'nodeType', '')) === 'mainSeq';

const NAME_CHILD_TN_LST = qname('p', 'childTnLst', NS.pml);
const NAME_CTN = qname('p', 'cTn', NS.pml);

const childPars = (el: XmlElement): XmlElement[] => {
  const childTnLst = firstChildElement(el, NAME_CHILD_TN_LST);
  if (childTnLst === null) return [];
  return childTnLst.children.filter(
    (c): c is XmlElement => c.kind === 'element' && isPml(c, 'par'),
  );
};

// The `<p:par>` wrappers nest click stop > group > effect, so one level down
// is `par > cTn > childTnLst > par`.
const innerPars = (par: XmlElement): XmlElement[] => {
  const cTn = firstChildElement(par, NAME_CTN);
  return cTn === null ? [] : childPars(cTn);
};

const appendPar = (par: XmlElement, child: XmlElement): boolean => {
  const cTn = firstChildElement(par, NAME_CTN);
  const childTnLst = cTn === null ? null : firstChildElement(cTn, NAME_CHILD_TN_LST);
  if (childTnLst === null) return false;
  childTnLst.children.push(child);
  return true;
};

const NAME_ST_COND_LST = qname('p', 'stCondLst', NS.pml);
const ATTR_DELAY = qname('', 'delay', '');
const ATTR_DUR = qname('', 'dur', '');

// Whole milliseconds only. `indefinite`, a missing value or anything that is
// not a plain integer means we cannot place a node on the timeline.
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
 * number of milliseconds, and treating one as the delay attribute alone would
 * place the node at a time PowerPoint never plays it.
 */
const offsetCond = (cTn: XmlElement): XmlElement | null => {
  const stCondLst = firstChildElement(cTn, NAME_ST_COND_LST);
  if (stCondLst === null) return null;
  const conds = stCondLst.children.filter(
    (c): c is XmlElement => c.kind === 'element' && isPml(c, 'cond'),
  );
  if (conds.length !== 1) return null;
  const cond = conds[0]!;
  if (getAttrValue(cond, qname('', 'evt', '')) !== null) return null;
  if (cond.children.some((c) => c.kind === 'element')) return null;
  return cond;
};

// A node with no start condition begins with its parent, which is offset 0.
const startOffsetMs = (cTn: XmlElement): number | null => {
  if (firstChildElement(cTn, NAME_ST_COND_LST) === null) return 0;
  const cond = offsetCond(cTn);
  return cond === null ? null : wholeMs(getAttrValue(cond, ATTR_DELAY));
};

const setStartOffset = (par: XmlElement, delayMs: number): boolean => {
  const cTn = firstChildElement(par, NAME_CTN);
  const cond = cTn === null ? null : offsetCond(cTn);
  if (cond === null) return false;
  cond.attrs = cond.attrs.map((a) =>
    a.name.namespaceURI === '' && a.name.localName === 'delay'
      ? { ...a, value: String(delayMs) }
      : a,
  );
  return true;
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
  const effectCTn = firstChildElement(effectPar, NAME_CTN);
  if (effectCTn === null || isRescaled(effectCTn)) return null;
  const start = startOffsetMs(effectCTn);
  if (start === null) return null;

  let longest: number | null = null;
  const walk = (el: XmlElement): boolean => {
    if (isPml(el, 'cBhvr')) {
      const cTn = firstChildElement(el, NAME_CTN);
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
 *
 * Sibling `<p:par>` nodes are parallel — they all begin when their parent
 * does — so `nodeType="afterEffect"` is a label for PowerPoint's UI, not a
 * dependency a player could honour. Placing the following group at this offset
 * is what actually makes it run afterwards.
 */
const groupEndMs = (groupPar: XmlElement): number | null => {
  const groupCTn = firstChildElement(groupPar, NAME_CTN);
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

// Merges a freshly-built single-effect timing into an existing `<p:timing>`,
// renumbering the new effect's cTn ids so they stay unique. Returns false when
// the existing tree has no structure we know how to extend (so the caller can
// avoid destroying it). This is what lets a second shape animate without wiping
// a template's pre-existing animations.
const mergeEffectInto = (
  existing: XmlElement,
  fresh: XmlElement,
  start: AnimationStartCondition,
): boolean => {
  const freshMainSeq = findDescendant(fresh, isMainSeqCTn);
  const freshChildTnLst = freshMainSeq
    ? firstChildElement(freshMainSeq, qname('p', 'childTnLst', NS.pml))
    : null;
  const newPar = freshChildTnLst
    ? freshChildTnLst.children.find((c): c is XmlElement => c.kind === 'element' && isPml(c, 'par'))
    : null;
  const freshBldP = findDescendant(fresh, (e) => isPml(e, 'bldP'));
  if (!newPar || !freshBldP) return false;

  // Group the build with its effect under a fresh grpId so PowerPoint renders
  // each shape's effect independently. Use max-existing-grpId + 1 (not a count)
  // because a template's authored build grpIds need not be the contiguous
  // 0..n-1 sequence — PowerPoint can leave gaps after a delete/reorder, and a
  // count would then collide with an existing group.
  const newGrpId = String(maxGrpId(existing) + 1);
  const effectCTn = findDescendant(
    newPar,
    (e) => getAttrValue(e, qname('', 'presetID', '')) !== null,
  );

  const existingMainSeq = findDescendant(existing, isMainSeqCTn);
  const existingMainSeqChildTnLst = existingMainSeq
    ? firstChildElement(existingMainSeq, qname('p', 'childTnLst', NS.pml))
    : null;
  if (existingMainSeqChildTnLst) {
    // The fresh tree's click-effect wrapper is the <p:par> under its own mainSeq
    // childTnLst. Lift it out and renumber its cTn ids past the existing max.
    const offset = maxCTnId(existing) - 2; // fresh effect ids start at 3
    if (offset > 0) shiftCTnIds(newPar, offset);

    // A click effect becomes its own stop. A with/after effect joins the stop
    // already there: `withPrevious` alongside the effects that run together,
    // `afterPrevious` as the next group inside the same stop, so both play off
    // the click that started their predecessor. With no stop to join — the
    // effect is the slide's first — the fresh wrapper is kept, and the builder
    // has already given it a zero delay so it runs as the slide appears.
    const stops = childPars(existingMainSeq!);
    const lastStop = stops.at(-1);
    const groups = lastStop === undefined ? [] : innerPars(lastStop);
    const lastGroup = groups.at(-1);

    if (start === 'click' || lastStop === undefined) {
      existingMainSeqChildTnLst.children.push(newPar);
    } else if (start === 'afterPrevious') {
      const group = innerPars(newPar)[0];
      if (group === undefined || lastGroup === undefined) return false;
      // The group is a sibling of the one before it, so it would otherwise
      // start at the same moment. Offsetting it to where that group finishes
      // is what makes "after previous" mean it; the caller's own delay stays
      // on the effect, counted from there.
      const previousEnd = groupEndMs(lastGroup);
      if (previousEnd === null) {
        throw new Error(
          'setShapeAnimation: cannot start an effect after one whose length this library cannot ' +
            'measure. The effect before it runs indefinitely, states no duration, repeats, or ' +
            'starts from another node rather than at a fixed offset. Use start: "click" or ' +
            '"withPrevious".',
        );
      }
      if (!setStartOffset(group, previousEnd) || !appendPar(lastStop, group)) return false;
    } else {
      const effectPar = innerPars(innerPars(newPar)[0] ?? newPar)[0];
      if (effectPar === undefined || lastGroup === undefined) return false;
      if (!appendPar(lastGroup, effectPar)) return false;
    }
  } else {
    // A slide that holds a video / audio clip but no animation yet has a root
    // with media nodes (and possibly interactive sequences) but no mainSeq.
    // Adopt the fresh tree's whole main sequence; PowerPoint keeps it first,
    // ahead of the interactive sequences and media nodes.
    const rootList = rootChildTnLst(existing);
    const freshRootList = rootChildTnLst(fresh);
    const freshSeq = freshRootList?.children.find(
      (c): c is XmlElement => c.kind === 'element' && isPml(c, 'seq'),
    );
    if (!rootList || !freshSeq) return false;
    shiftCTnIds(freshSeq, maxCTnId(existing) - 1); // fresh mainSeq ids start at 2
    rootList.children.unshift(freshSeq);
  }

  if (effectCTn) setGrpId(effectCTn, newGrpId);
  setGrpId(freshBldP, newGrpId);
  const existingBldLst = findDescendant(existing, (e) => isPml(e, 'bldLst'));
  if (existingBldLst) existingBldLst.children.push(freshBldP);
  else insertBldLst(existing, elem(qname('p', 'bldLst', NS.pml), { children: [freshBldP] }));
  return true;
};

// CT_SlideTiming orders its children tnLst, bldLst, extLst.
const insertBldLst = (timing: XmlElement, bldLst: XmlElement): void => {
  const extLst = firstChildElement(timing, qname('p', 'extLst', NS.pml));
  if (extLst === null) timing.children.push(bldLst);
  else timing.children.splice(timing.children.indexOf(extLst), 0, bldLst);
};

/**
 * Adds a single animation effect to the given shape.
 *
 * The effect is *merged* into any existing `<p:timing>` on the slide rather
 * than replacing it: animating a second shape (or re-running on a template that
 * already has authored animations) preserves the existing effects, with cTn ids
 * renumbered to stay unique. To clear every animation first, call
 * `clearSlideAnimations`.
 *
 * Supported `effect` tokens:
 *
 *   - `'fadeIn'`   entrance fade
 *   - `'fadeOut'`  exit fade
 *   - `'appear'`   instant entrance
 *   - `'disappear'` instant exit
 *
 * `durationMs` defaults to 500ms (fades only — `appear`/`disappear`
 * are instantaneous).
 *
 * `start` decides where the effect lands. The default `'click'` gives it a
 * click stop of its own, so the viewer sees it on the next click.
 * `'withPrevious'` and `'afterPrevious'` join the stop the last effect is in,
 * running alongside it or once it has finished — both off the same click. As
 * the slide's first effect neither has a predecessor, so both run when the
 * slide appears. `delayMs` waits that long once the start condition is met.
 *
 * `'afterPrevious'` throws when the slide's timing does not say when the effect
 * before it ends — it runs indefinitely, states no duration, repeats, or starts
 * from another node rather than at a fixed offset. The slide is left untouched;
 * `'click'` and `'withPrevious'` need no such measurement.
 *
 * Reading the result back, including the order and start condition of every
 * effect on the slide, is `getSlideAnimations`.
 */
export const setShapeAnimation = (shape: SlideShapeData, opts: AnimationOptions): void => {
  const slide = shape[SHAPE_SLIDE];
  const spid = shape[SHAPE_SNAPSHOT].id;
  const fresh = buildSingleEffectTiming(spid, opts);
  const existing = findTiming(slide);
  if (existing === null) {
    insertTimingAtEnd(slide, fresh);
  } else if (!mergeEffectInto(existing, fresh, opts.start ?? 'click')) {
    // A timing tree we don't know how to extend (no mainSeq). Leave it intact
    // rather than silently destroying authored animations.
    throw new Error(
      'setShapeAnimation: the slide already has an animation timing tree this single-effect ' +
        'API cannot safely extend. Call clearSlideAnimations(slide) first to reset it.',
    );
  }
  commitSlideData(slide);
  refreshSlideData(slide);
};

/**
 * Returns the animation effect bound to this shape via the slide's
 * `<p:timing>` tree, or `null` when the shape has none. A shape carrying
 * several effects reports the first one in document order. Presets outside
 * the four `AnimationEffect` tokens are reported as `null` rather than
 * guessed at.
 */
export const getShapeAnimation = (shape: SlideShapeData): AnimationEffect | null =>
  firstEffectByShape(shape[SHAPE_SLIDE]).get(shape[SHAPE_SNAPSHOT].id) ?? null;

// Only a time node that names a preset was ever a candidate for the
// shape-level readers, so an effect written without `presetID` / `presetClass`
// does not shadow the recognised effect behind it. `getSlideAnimations` lists
// both; this narrower rule belongs to the two shape-level readers alone.
const namesAPreset = (cTn: XmlElement): boolean => {
  const presetId = getAttrValue(cTn, qname('', 'presetID', ''));
  const presetClass = getAttrValue(cTn, qname('', 'presetClass', ''));
  return presetId !== null && presetId !== '' && presetClass !== null && presetClass !== '';
};

// PowerPoint needs a `<p:bldLst><p:bldP spid="...">` entry for an effect to
// render, so a shape without one has no animation as far as the read API is
// concerned. Built once per call and shared by the two shape-level readers,
// which would otherwise re-walk the whole tree for every shape on the slide.
const firstEffectByShape = (slide: SlideData): Map<number, AnimationEffect | null> => {
  const out = new Map<number, AnimationEffect | null>();
  const timing = findSlideTimingElement(slide);
  if (timing === null) return out;
  const bldLst = firstChildElement(timing, qname('p', 'bldLst', NS.pml));
  if (bldLst === null) return out;
  const built = new Set(
    allChildElements(bldLst, qname('p', 'bldP', NS.pml)).map((b) =>
      getAttrValue(b, qname('', 'spid', '')),
    ),
  );
  for (const { step, cTn } of readSlideTiming(slide)) {
    if (!namesAPreset(cTn)) continue;
    for (const shapeId of step.targetShapeIds) {
      if (!out.has(shapeId) && built.has(String(shapeId))) out.set(shapeId, step.effect);
    }
  }
  return out;
};

/**
 * Every animation effect on the slide, in document order — the click order a
 * viewer sees. Effects this library cannot author are still listed, with
 * `editable: false`, so a caller can show them instead of losing them.
 */
export const getSlideAnimations = (slide: SlideData): readonly SlideAnimationStep[] =>
  readSlideTiming(slide).map((node) => node.step);

/**
 * Returns every shape on the slide that has an authored animation
 * effect (i.e. `getShapeAnimation(shape)` is not `null`). Pair to
 * `slideHasAnimations`. Useful for audit reports — "which shapes on
 * this slide actually animate?" before exporting to a video pipeline
 * that doesn't honor PowerPoint's timing tree.
 */
export const findShapesWithAnimation = (slide: SlideData): ReadonlyArray<SlideShapeData> => {
  const effects = firstEffectByShape(slide);
  return slide[SLIDE_SHAPES].filter((shape) => effects.get(shape[SHAPE_SNAPSHOT].id) != null);
};

/** Removes every animation on the slide, keeping its media play controls. */
export const clearSlideAnimations = (slide: SlideData): void => {
  // Media time nodes are what give a video / audio clip its play controls,
  // not animations — dropping them with the rest would silently break the
  // slide's clips.
  const existing = findTiming(slide);
  const mediaNodes = existing ? mediaTimingNodes(existing) : [];
  if (existing !== null && mediaNodes.length > 0) {
    const root = slide[SLIDE_DOCUMENT].root;
    root.children[root.children.indexOf(existing)] = buildTimingRoot(mediaNodes);
  } else {
    removeExistingTiming(slide);
  }
  commitSlideData(slide);
  refreshSlideData(slide);
};
