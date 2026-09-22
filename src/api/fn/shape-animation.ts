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
  attr,
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
  buildEntryKey,
  cTnIdsUnder,
  delayOf,
  isPlainWrapper,
  readBuildEntries,
  remapSpids,
  startOfEffect,
} from './_animation-layout.ts';
import {
  type AnimationSequenceKind,
  type AnimationStart,
  type AnimationValueAfterEnd,
  type AnimationTarget,
  type SlideAnimationStep,
  findSlideTimingElement,
  groupEndMs,
  readSlideTiming,
  readTimingSteps,
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
export type { AnimationValueAfterEnd };

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

/**
 * Gives every `<p:cTn>` under `el` the next id the cursor has, in document
 * order.
 *
 * A freshly built effect numbers its nodes from 3, so shifting them past the
 * tree would do; a copied subtree carries whatever ids the file it came from
 * used, and those can land anywhere — on top of an id already in the target
 * included. Renumbering outright is the one rule that holds for both.
 */
const renumberCTnIds = (el: XmlElement, cursor: MergeCursor): void => {
  const walk = (e: XmlElement): void => {
    if (isPml(e, 'cTn')) {
      const previous = getAttrValue(e, ATTR_ID_FN);
      const id = String(++cursor.maxCTnId);
      e.attrs =
        previous === null
          ? [attr(ATTR_ID_FN, id), ...e.attrs]
          : e.attrs.map((a) =>
              a.name.namespaceURI === '' && a.name.localName === 'id' ? { ...a, value: id } : a,
            );
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
  renumberCTnIds(effect.par, cursor);

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
    const fresh = buildSingleEffectTiming(spid, opts, { paragraph });
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

// ---------------------------------------------------------------------------
// Copying a shape's animations along with the shape.
//
// Each effect is copied as the subtree it is, not rebuilt from what the read
// model makes of it: a build whose second paragraph was deleted stays deleted,
// a paragraph retimed on its own keeps its own timing, and a preset this
// library cannot name still comes across. Only the shape ids, the cTn ids and
// the build groups are rewritten.
//
// The whole copy — the offsets an `afterPrevious` effect needs measured
// included — is assembled before the caller has added a shape, a relationship
// or a part, so a refusal costs the target nothing.

const COPY = 'copyShape: ';

const ATTR_GRP_ID_FN = qname('', 'grpId', '');
const ATTR_VAL_FN = qname('', 'val', '');
const ATTR_SPID_FN = qname('', 'spid', '');

/**
 * The `<p:par>` each time node hangs from, and the `<p:par>` that one hangs
 * from in turn — the ancestry a copied effect is checked against, collected in
 * one walk rather than searched upwards per effect.
 */
interface Ancestry {
  readonly parOf: ReadonlyMap<XmlElement, XmlElement>;
  readonly parentOf: ReadonlyMap<XmlElement, XmlElement>;
}

const ancestryOf = (timing: XmlElement): Ancestry => {
  const parOf = new Map<XmlElement, XmlElement>();
  const parentOf = new Map<XmlElement, XmlElement>();
  const walk = (el: XmlElement, enclosing: XmlElement | null): void => {
    let par = enclosing;
    if (isPml(el, 'par')) {
      const cTn = firstChildElement(el, NAME_CTN);
      if (cTn !== null) parOf.set(cTn, el);
      if (enclosing !== null) parentOf.set(el, enclosing);
      par = el;
    }
    for (const c of el.children) if (c.kind === 'element') walk(c, par);
  };
  walk(timing, null);
  return { parOf, parentOf };
};

/** One effect of the source slide, and what the copy needs to know about it. */
interface CopiedEffect {
  /** The source's effect `<p:par>` — the subtree that is cloned. */
  readonly par: XmlElement;
  readonly start: AnimationStartCondition;
  /** The shapes it drives, in document order; all of them are being copied. */
  readonly spids: readonly string[];
  readonly grpId: string | null;
}

/**
 * The effects to copy, in the click order they play in. Taken from the read
 * model's own effect nodes rather than by looking for `<p:par>` wrappers whose
 * targets all happen to be copied: on a slide where the copied shape is the
 * only animated one, every wrapper up to the root passes that test, and the
 * root is not an effect.
 */
const effectsToCopy = (source: XmlElement, copied: ReadonlySet<string>): CopiedEffect[] => {
  const { parOf, parentOf } = ancestryOf(source);
  const checked = new Map<XmlElement, ReadonlySet<XmlElement> | null>();
  const out: CopiedEffect[] = [];
  for (const node of readTimingSteps(source)) {
    const spids = node.step.targetShapeIds.map(String);
    if (!spids.some((id) => copied.has(id))) continue;
    if (!spids.every((id) => copied.has(id))) {
      throw new Error(
        `${COPY}this shape shares an animation with a shape that is not being copied, so the copy ` +
          'cannot be given one of its own. Copy the shapes together, or split the effect first.',
      );
    }
    if (node.step.sequence !== 'mainSeq') {
      throw new Error(
        `${COPY}this shape is animated by an interactive sequence, which is bound to the click ` +
          'that triggers it and cannot be reproduced for the copy. Remove that animation first.',
      );
    }
    const par = parOf.get(node.cTn);
    const start = startOfEffect(node.cTn);
    if (par === undefined || start === null) {
      throw new Error(
        `${COPY}this shape has an animation whose start condition is not one this library models, ` +
          'so the copy could not be given the same one. Remove it from the original first.',
      );
    }
    assertWrappersAreOurs(par, parentOf, checked);
    out.push({ par, start, spids, grpId: getAttrValue(node.cTn, ATTR_GRP_ID_FN) });
  }
  return out;
};

/**
 * The groups of one click stop that this library could have written, or `null`
 * when the stop itself could not be.
 *
 * The copy is given the stop and group this library writes, with the start
 * read off the effect's own `nodeType` and the offset recomputed against the
 * target slide. Anything else those wrappers carried — a repeat, an event
 * condition, an extension, or a wait this library would not have written —
 * would be dropped on the way, so such an effect is refused instead.
 *
 * Read once per stop, walking its groups in order: the delay this library
 * would give a group is the end of the group before it, so checking them one
 * at a time from the front costs a single pass, where asking the question per
 * effect would re-measure the same stop once per paragraph of a build.
 */
const ourGroupsIn = (stop: XmlElement): ReadonlySet<XmlElement> | null => {
  if (!isPlainWrapper(stop)) return null;
  const groups = innerPars(stop);

  // A stop waits for the viewer when the first effect in it is a click
  // effect, and opens with the slide when that effect runs with or after a
  // predecessor it does not have. Copying one written the other way round
  // would turn an automatic start into a click, or a click into one.
  const first = groups[0];
  const firstEffect = first === undefined ? undefined : innerPars(first)[0];
  const firstCTn = firstEffect === undefined ? null : firstChildElement(firstEffect, NAME_CTN);
  const firstStart = firstCTn === null ? null : startOfEffect(firstCTn);
  if (firstStart === null) return null;
  if (delayOf(stop) !== (firstStart === 'click' ? 'indefinite' : '0')) return null;

  const ours = new Set<XmlElement>();
  let expected: number | null = 0;
  for (const group of groups) {
    if (expected !== null && isPlainWrapper(group) && delayOf(group) === String(expected)) {
      ours.add(group);
    }
    expected = groupEndMs(group);
  }
  return ours;
};

/**
 * Refuses an effect whose click stop or group says more than when it starts.
 *
 * Only the ancestors of the effects being copied are read. An effect on some
 * other shape may be wrapped however its author liked without standing in the
 * way of this copy.
 */
const assertWrappersAreOurs = (
  effectPar: XmlElement,
  parentOf: ReadonlyMap<XmlElement, XmlElement>,
  checked: Map<XmlElement, ReadonlySet<XmlElement> | null>,
): void => {
  const group = parentOf.get(effectPar);
  const stop = group === undefined ? undefined : parentOf.get(group);
  if (group === undefined || stop === undefined) throw wrapperRefusal();
  let ours = checked.get(stop);
  if (ours === undefined) {
    ours = ourGroupsIn(stop);
    checked.set(stop, ours);
  }
  if (ours === null || !ours.has(group)) throw wrapperRefusal();
};

const wrapperRefusal = (): Error =>
  new Error(
    `${COPY}this shape's animation is held in a click stop or group that does more than say when ` +
      'it starts, and the copy would lose that. Remove it from the original first, or give the ' +
      'copy an animation of its own with setShapeAnimation.',
  );

/** Every `<p:cTn id>` at or under `el`, in document order — renumbering's key. */
const cTnIdsInOrder = (el: XmlElement): (string | null)[] => {
  const out: (string | null)[] = [];
  const walk = (e: XmlElement): void => {
    if (isPml(e, 'cTn')) out.push(getAttrValue(e, ATTR_ID_FN));
    for (const c of e.children) if (c.kind === 'element') walk(c);
  };
  walk(el);
  return out;
};

/** Every `<p:tn val>` at or under `el` — the time nodes it waits on. */
const timeNodeRefs = (el: XmlElement, into: Set<string>): void => {
  if (isPml(el, 'tn')) {
    const val = getAttrValue(el, ATTR_VAL_FN);
    if (val !== null) into.add(val);
  }
  for (const c of el.children) if (c.kind === 'element') timeNodeRefs(c, into);
};

/** Rewrites those references to the ids the copies were renumbered to. */
const remapTimeNodeRefs = (el: XmlElement, remap: ReadonlyMap<string, string>): void => {
  if (isPml(el, 'tn')) {
    el.attrs = el.attrs.map((a) =>
      a.name.namespaceURI === '' && a.name.localName === 'val'
        ? { ...a, value: remap.get(a.value) ?? a.value }
        : a,
    );
  }
  for (const c of el.children) if (c.kind === 'element') remapTimeNodeRefs(c, remap);
};

/**
 * Refuses an effect that waits on a time node outside the copied set. One that
 * waits on another copied effect is fine — both are renumbered through the
 * same map, so the copy waits on its own predecessor — but one that waits on
 * an effect staying behind would follow the original, and stop following
 * anything the moment the original is retimed.
 */
const assertReferencesStayInside = (effects: readonly CopiedEffect[]): void => {
  const inside = new Set<string>();
  for (const effect of effects) for (const id of cTnIdsUnder(effect.par)) inside.add(id);
  const refs = new Set<string>();
  for (const effect of effects) timeNodeRefs(effect.par, refs);
  const outside = [...refs].find((id) => !inside.has(id));
  if (outside !== undefined) {
    throw new Error(
      `${COPY}this shape has an animation timed against another element of the slide ` +
        `(<p:tn val="${outside}">), and the copy would either still wait on the original or wait ` +
        'on nothing. Remove it from the original first.',
    );
  }
};

/**
 * A `<p:timing>` with a main sequence and nothing in it, for a target slide
 * that has no timing at all. Every effect then merges the way it would on a
 * slide that already animates something, rather than the first one taking a
 * path of its own.
 */
const emptyMainSeqTiming = (): XmlElement => {
  const timing = buildSingleEffectTiming(1, { effect: 'appear' });
  const mainSeq = findDescendant(timing, isMainSeqCTn);
  const childTnLst = mainSeq === null ? null : firstChildElement(mainSeq, NAME_CHILD_TN_LST);
  if (childTnLst === null) throw new Error(`${COPY}post-condition failed: no main sequence`);
  childTnLst.children = [];
  timing.children = timing.children.filter((c) => !(c.kind === 'element' && isPml(c, 'bldLst')));
  return timing;
};

/**
 * Puts an empty main sequence ahead of the sequences a timing already has, for
 * a target slide whose timing holds only media or interactive nodes.
 * PowerPoint keeps the main sequence first.
 */
const adoptEmptyMainSeq = (timing: XmlElement): MergeCursor | null => {
  const rootList = rootChildTnLst(timing);
  const seq = rootChildTnLst(emptyMainSeqTiming())?.children.find(
    (c): c is XmlElement => c.kind === 'element' && isPml(c, 'seq'),
  );
  if (!rootList || !seq) return null;
  shiftCTnIds(seq, maxCTnId(timing) - 1); // a fresh mainSeq is cTn id 2
  rootList.children.unshift(seq);
  return openMergeCursor(timing);
};

/** The builder's click stop and group wrapper, with a cloned effect inside. */
const wrapClonedEffect = (
  spid: string,
  start: AnimationStartCondition,
  clone: XmlElement,
  bldP: XmlElement,
): FreshEffect | null => {
  const shell = openFreshEffect(buildSingleEffectTiming(Number(spid), { effect: 'appear', start }));
  const group = shell === null ? undefined : innerPars(shell.par)[0];
  const groupCTn = group === undefined ? null : firstChildElement(group, NAME_CTN);
  const childTnLst = groupCTn === null ? null : firstChildElement(groupCTn, NAME_CHILD_TN_LST);
  const cloneCTn = firstChildElement(clone, NAME_CTN);
  if (shell === null || childTnLst === null || cloneCTn === null) return null;
  childTnLst.children = [clone];
  return { par: shell.par, bldP, cTn: cloneCTn };
};

/** A `<p:bldP>` for a copied shape whose original had none. */
const freshBuildEntry = (spid: string): XmlElement =>
  elem(qname('p', 'bldP', NS.pml), {
    attrs: [attr(ATTR_SPID_FN, spid), attr(ATTR_GRP_ID_FN, '0')],
  });

/** The build entries a newly opened group needs — one per shape it animates. */
const buildEntriesFor = (
  effect: CopiedEffect,
  idMap: ReadonlyMap<string, string>,
  sourceBuilds: ReadonlyMap<string, XmlElement>,
): XmlElement[] =>
  effect.spids.map((spid) => {
    const copied = idMap.get(spid)!;
    const source = sourceBuilds.get(buildEntryKey(spid, effect.grpId));
    if (source === undefined) return freshBuildEntry(copied);
    // The original's entry carries `build="p"` and `bldLvl`, which is what
    // makes a copied paragraph build still reveal a paragraph at a time.
    const entry = cloneElement(source);
    entry.attrs = entry.attrs.map((a) =>
      a.name.namespaceURI === '' && a.name.localName === 'spid' ? { ...a, value: copied } : a,
    );
    return entry;
  });

const NO_ANIMATION_COPY = (): void => {};

/**
 * Works out what `targetSlide`'s timing has to become for the copies of the
 * shapes in `idMap` to animate the way their originals do, and returns the
 * step that installs it. Throws instead when the animations cannot be
 * reproduced — while the caller has still changed nothing, which is the point
 * of splitting the work in two.
 *
 * A trigger is not inherited: an interactive sequence the original starts
 * stays bound to the original.
 *
 * @internal
 */
export const planAnimationCopy = (
  sourceSlide: SlideData,
  targetSlide: SlideData,
  idMap: ReadonlyMap<string, string>,
): (() => void) => {
  const source = findSlideTimingElement(sourceSlide);
  if (source === null || idMap.size === 0) return NO_ANIMATION_COPY;
  const effects = effectsToCopy(source, new Set(idMap.keys()));
  if (effects.length === 0) return NO_ANIMATION_COPY;
  assertReferencesStayInside(effects);

  const sourceBuilds = readBuildEntries(source);
  const existing = findTiming(targetSlide);
  const timing = existing === null ? emptyMainSeqTiming() : cloneElement(existing);
  const cursor = openMergeCursor(timing) ?? adoptEmptyMainSeq(timing);
  if (cursor === null) {
    throw new Error(
      `${COPY}the target slide has an animation timing tree this library cannot safely extend, so ` +
        "the copy's animations have nowhere to go. Call clearSlideAnimations on it first.",
    );
  }

  // A new group takes max-existing + 1 rather than a count: authored grpIds
  // need not be the contiguous 0..n sequence, and a count would collide.
  let nextGrpId = maxGrpId(timing) + 1;
  const groups = new Map<string, string>();
  const renumbered = new Map<string, string>();
  const clones: XmlElement[] = [];

  for (const effect of effects) {
    const clone = cloneElement(effect.par);
    remapSpids(clone, idMap);
    clones.push(clone);
    // Read before the merge renumbers them, and again after. Only the copied
    // effect's own nodes go into the reference map: the click stop and group
    // the copy is given are this library's own, numbered from the same
    // counter, and a source effect that happened to use one of their numbers
    // would otherwise take their new id and leave `<p:tn>` pointing at a
    // wrapper — one that `withPrevious` does not even keep.
    const wasNumbered = cTnIdsInOrder(clone);

    // One group per source group, so the paragraphs of one build still share a
    // single entry and animate as one build on the copy too.
    const key = `${effect.spids.join(',')}|${effect.grpId ?? ''}`;
    const known = groups.get(key);
    const grpId = known ?? String(nextGrpId++);
    const entries = known === undefined ? buildEntriesFor(effect, idMap, sourceBuilds) : [];
    const spid = idMap.get(effect.spids[0]!)!;
    const wrapped = wrapClonedEffect(
      spid,
      effect.start,
      clone,
      entries[0] ?? freshBuildEntry(spid),
    );
    const merged =
      wrapped !== null &&
      mergeEffectInto(cursor, wrapped, effect.start, { grpId, addBuild: known === undefined });
    if (!merged) {
      throw new Error(
        `${COPY}the target slide has an animation timing tree this library cannot safely extend, ` +
          "so the copy's animations have nowhere to go. Call clearSlideAnimations on it first.",
      );
    }
    const nowNumbered = cTnIdsInOrder(clone);
    for (const [i, previous] of wasNumbered.entries()) {
      const id = nowNumbered[i];
      if (previous !== null && id != null) renumbered.set(previous, id);
    }
    // A composite effect drives more than one shape, and PowerPoint wants an
    // entry per shape under the group they share.
    for (const extra of entries.slice(1)) {
      setGrpId(extra, grpId);
      cursor.bldLst = addBuildEntry(cursor.timing, extra, cursor.bldLst);
    }
    groups.set(key, grpId);
  }

  // Only now is every new id known, so an effect that waited on another copied
  // one can be pointed at that one's copy.
  for (const clone of clones) remapTimeNodeRefs(clone, renumbered);

  return () => {
    const children = targetSlide[SLIDE_DOCUMENT].root.children;
    if (existing === null) insertTimingAtEnd(targetSlide, timing);
    else children[children.indexOf(existing)] = timing;
  };
};
