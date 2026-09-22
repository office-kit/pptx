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
  cloneElement,
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
  groupEndMs,
  readSlideTiming,
  setGroupStartOffset,
} from './_animation-timing.ts';
import { commitSlideData, refreshSlideData } from './_helpers.ts';
import { getShapeParagraphCount } from './shape-runs.ts';
import { maxCTnId, mediaTimingNodes, rootChildTnLst } from './_media-timing.ts';
// ---------------------------------------------------------------------------
// Animations (one effect per call).
//
// Current scope: each call adds one entrance or exit effect — or, for a
// by-paragraph build, one per paragraph — and merges it into whatever timing
// tree the slide already has, so a slide can carry several effects across
// several shapes and click stops. Emphasis and motion presets, and editing or
// reordering an effect that is already there, are not modelled yet.

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

/**
 * The build group a merged effect joins. PowerPoint ties an effect to its
 * `<p:bldP>` through `grpId`, and a by-paragraph build is several effects
 * sharing one entry, so only the first of them carries `addBuild`.
 */
interface BuildGroup {
  readonly grpId: string;
  readonly addBuild: boolean;
}

/**
 * Where the next effect goes, and the ids it must not collide with.
 *
 * One call can add several effects — a by-paragraph build adds one per
 * paragraph — and each of them lands at the end of a tree the call before it
 * just grew. Re-deriving the insertion point from the whole tree every time
 * would make a build of n paragraphs cost O(n²), so it is derived once and
 * moved forward as effects land.
 */
interface MergeCursor {
  readonly timing: XmlElement;
  readonly mainSeqChildTnLst: XmlElement;
  /** Largest `<p:cTn id>` in the tree; a fresh effect is renumbered past it. */
  maxCTnId: number;
  /** The last click stop and the last group in it — what with/after join. */
  lastStop: XmlElement | undefined;
  lastGroup: XmlElement | undefined;
  bldLst: XmlElement | null;
}

const openMergeCursor = (timing: XmlElement): MergeCursor | null => {
  const mainSeq = findDescendant(timing, isMainSeqCTn);
  const childTnLst = mainSeq === null ? null : firstChildElement(mainSeq, NAME_CHILD_TN_LST);
  if (mainSeq === null || childTnLst === null) return null;
  const lastStop = childPars(mainSeq).at(-1);
  return {
    timing,
    mainSeqChildTnLst: childTnLst,
    maxCTnId: maxCTnId(timing),
    lastStop,
    lastGroup: lastStop === undefined ? undefined : innerPars(lastStop).at(-1),
    bldLst: findDescendant(timing, (e) => isPml(e, 'bldLst')),
  };
};

// CT_SlideTiming orders its children tnLst, bldLst, extLst.
const insertBldLst = (timing: XmlElement, bldLst: XmlElement): void => {
  const extLst = firstChildElement(timing, qname('p', 'extLst', NS.pml));
  if (extLst === null) timing.children.push(bldLst);
  else timing.children.splice(timing.children.indexOf(extLst), 0, bldLst);
};

// PowerPoint renders an effect only when a `<p:bldP>` names its shape and
// group, so every group that opens brings its entry with it.
const addBuildEntry = (
  timing: XmlElement,
  bldP: XmlElement,
  bldLst: XmlElement | null,
): XmlElement => {
  if (bldLst !== null) {
    bldLst.children.push(bldP);
    return bldLst;
  }
  const fresh = elem(qname('p', 'bldLst', NS.pml), { children: [bldP] });
  insertBldLst(timing, fresh);
  return fresh;
};

/** The pieces of a freshly-built single-effect tree the merge paths need. */
interface FreshEffect {
  readonly par: XmlElement;
  readonly bldP: XmlElement;
  readonly cTn: XmlElement | null;
}

const openFreshEffect = (fresh: XmlElement): FreshEffect | null => {
  const mainSeq = findDescendant(fresh, isMainSeqCTn);
  const childTnLst = mainSeq === null ? null : firstChildElement(mainSeq, NAME_CHILD_TN_LST);
  const par = childTnLst?.children.find(
    (c): c is XmlElement => c.kind === 'element' && isPml(c, 'par'),
  );
  const bldP = findDescendant(fresh, (e) => isPml(e, 'bldP'));
  if (par === undefined || bldP === null) return null;
  return {
    par,
    bldP,
    cTn: findDescendant(par, (e) => getAttrValue(e, qname('', 'presetID', '')) !== null),
  };
};

/**
 * Adopts a fresh tree's whole main sequence into a timing that has none: a
 * slide holding a video or audio clip has a root with media nodes, and
 * possibly interactive sequences, but no `mainSeq`. PowerPoint keeps the main
 * sequence first, ahead of both.
 *
 * Returns the cursor for the sequence it just added, so the effects after this
 * one merge into it like they would on any other slide.
 */
const adoptMainSeq = (
  timing: XmlElement,
  fresh: XmlElement,
  effect: FreshEffect,
  group: BuildGroup,
): MergeCursor | null => {
  const rootList = rootChildTnLst(timing);
  const freshSeq = rootChildTnLst(fresh)?.children.find(
    (c): c is XmlElement => c.kind === 'element' && isPml(c, 'seq'),
  );
  if (!rootList || !freshSeq) return null;
  shiftCTnIds(freshSeq, maxCTnId(timing) - 1); // fresh mainSeq ids start at 2
  rootList.children.unshift(freshSeq);

  if (effect.cTn) setGrpId(effect.cTn, group.grpId);
  if (group.addBuild) {
    setGrpId(effect.bldP, group.grpId);
    addBuildEntry(
      timing,
      effect.bldP,
      findDescendant(timing, (e) => isPml(e, 'bldLst')),
    );
  }
  return openMergeCursor(timing);
};

/**
 * Places one freshly-built effect at the cursor, renumbering its cTn ids past
 * everything already in the tree, and moves the cursor to it. Returns false
 * for a tree whose shape we do not recognise, so the caller can abandon the
 * draft rather than half-writing it.
 */
const mergeEffectInto = (
  cursor: MergeCursor,
  effect: FreshEffect,
  start: AnimationStartCondition,
  group: BuildGroup,
): boolean => {
  const offset = cursor.maxCTnId - 2; // fresh effect ids start at 3
  if (offset > 0) shiftCTnIds(effect.par, offset);
  cursor.maxCTnId = Math.max(cursor.maxCTnId, maxCTnId(effect.par));

  // A click effect becomes its own stop. A with/after effect joins the stop
  // already there: `withPrevious` alongside the effects that run together,
  // `afterPrevious` as the next group inside the same stop, so both play off
  // the click that started their predecessor. With no stop to join — the
  // effect is the slide's first — the fresh wrapper is kept, and the builder
  // has already given it a zero delay so it runs as the slide appears.
  if (start === 'click' || cursor.lastStop === undefined) {
    cursor.mainSeqChildTnLst.children.push(effect.par);
    cursor.lastStop = effect.par;
    cursor.lastGroup = innerPars(effect.par).at(-1);
  } else if (start === 'afterPrevious') {
    const group_ = innerPars(effect.par)[0];
    if (group_ === undefined || cursor.lastGroup === undefined) return false;
    // The group is a sibling of the one before it, so it would otherwise
    // start at the same moment. Offsetting it to where that group finishes
    // is what makes "after previous" mean it; the caller's own delay stays
    // on the effect, counted from there.
    const previousEnd = groupEndMs(cursor.lastGroup);
    if (previousEnd === null) {
      throw new Error(
        'setShapeAnimation: cannot start an effect after one whose length this library cannot ' +
          'measure. The effect before it runs indefinitely, states no duration, repeats, or ' +
          'starts from another node rather than at a fixed offset. Use start: "click" or ' +
          '"withPrevious".',
      );
    }
    if (!setGroupStartOffset(group_, previousEnd) || !appendPar(cursor.lastStop, group_)) {
      return false;
    }
    cursor.lastGroup = group_;
  } else {
    const effectPar = innerPars(innerPars(effect.par)[0] ?? effect.par)[0];
    if (effectPar === undefined || cursor.lastGroup === undefined) return false;
    if (!appendPar(cursor.lastGroup, effectPar)) return false;
  }

  if (effect.cTn) setGrpId(effect.cTn, group.grpId);
  if (group.addBuild) {
    setGrpId(effect.bldP, group.grpId);
    cursor.bldLst = addBuildEntry(cursor.timing, effect.bldP, cursor.bldLst);
  }
  return true;
};

/**
 * The timing this call should leave on the slide: a copy of `existing` with
 * every effect added, or the builder's own tree when the slide has none.
 * `null` for an existing tree we cannot extend.
 *
 * Working on a copy is what makes a call all-or-nothing. A by-paragraph build
 * adds one effect per paragraph, and a refusal part-way must leave the slide
 * as it was rather than stranding it with some of the paragraphs animated.
 */
const timingWithEffects = (
  existing: XmlElement | null,
  spid: number,
  opts: AnimationOptions,
  targets: readonly (number | null)[],
): XmlElement | null => {
  const start = opts.start ?? 'click';
  let timing = existing === null ? null : cloneElement(existing);
  let cursor = timing === null ? null : openMergeCursor(timing);
  // Every paragraph of one build joins the group the first of them opened. A
  // fresh group takes max-existing-grpId + 1 (not a count) because a template's
  // authored build grpIds need not be the contiguous 0..n-1 sequence —
  // PowerPoint can leave gaps after a delete or reorder, and a count would then
  // collide with an existing group.
  let buildGrpId: string | null = null;

  for (const paragraph of targets) {
    const fresh = buildSingleEffectTiming(spid, opts, paragraph);
    if (timing === null) {
      // No timing on the slide at all, so the builder's standalone tree is the
      // draft — effect, click stop and build entry already in place, with the
      // group it numbers 0.
      timing = fresh;
      cursor = openMergeCursor(timing);
      if (cursor === null) return null;
      buildGrpId = '0';
      continue;
    }
    const effect = openFreshEffect(fresh);
    if (effect === null) return null;
    const group: BuildGroup = {
      grpId: buildGrpId ?? String(maxGrpId(timing) + 1),
      addBuild: buildGrpId === null,
    };
    if (cursor === null) {
      cursor = adoptMainSeq(timing, fresh, effect, group);
      if (cursor === null) return null;
    } else if (!mergeEffectInto(cursor, effect, start, group)) {
      return null;
    }
    buildGrpId = group.grpId;
  }
  return timing;
};

/**
 * The paragraphs each effect of this call targets, in order. `null` is the
 * whole shape — one effect, the way every call worked before builds existed.
 */
const effectTargets = (shape: SlideShapeData, opts: AnimationOptions): (number | null)[] => {
  if (opts.byParagraph !== true) return [null];
  const count = getShapeParagraphCount(shape);
  if (count === 0) {
    throw new Error(
      'setShapeAnimation: byParagraph needs a shape with text. This shape has no paragraphs to ' +
        'build, so animate it as a whole instead.',
    );
  }
  return Array.from({ length: count }, (_, i) => i);
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
 * `byParagraph` reveals the shape's text one paragraph at a time rather than
 * animating the shape as a whole: one effect per paragraph, each with the same
 * `start`, so the default `'click'` advances a paragraph per click. They share
 * a single `<p:bldP build="p">`, which is how PowerPoint and Google Slides both
 * present the build as one animation. `getSlideAnimations` reports each
 * paragraph as its own step, targeting a paragraph range.
 *
 * Reading the result back, including the order and start condition of every
 * effect on the slide, is `getSlideAnimations`.
 */
export const setShapeAnimation = (shape: SlideShapeData, opts: AnimationOptions): void => {
  const slide = shape[SHAPE_SLIDE];
  const spid = shape[SHAPE_SNAPSHOT].id;
  const existing = findTiming(slide);
  const draft = timingWithEffects(existing, spid, opts, effectTargets(shape, opts));
  if (draft === null) {
    // A timing tree we don't know how to extend. Leave it intact rather than
    // silently destroying authored animations.
    throw new Error(
      'setShapeAnimation: the slide already has an animation timing tree this single-effect ' +
        'API cannot safely extend. Call clearSlideAnimations(slide) first to reset it.',
    );
  }

  const children = slide[SLIDE_DOCUMENT].root.children;
  if (existing === null) insertTimingAtEnd(slide, draft);
  else children[children.indexOf(existing)] = draft;
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
