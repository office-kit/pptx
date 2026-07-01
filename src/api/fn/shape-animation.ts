// Slide animations.

import {
  type AnimationEffect,
  type AnimationOptions,
  buildSingleEffectTiming,
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
import { commitSlideData, refreshSlideData } from './_helpers.ts';
// ---------------------------------------------------------------------------
// Animations (single-effect, click-triggered).
//
// v1 scope: exactly one effect per slide, click-triggered, entrance or
// exit preset family. The plan calls this the curated subset; full
// multi-effect timing-tree authoring is post-1.0.

export type { AnimationEffect, AnimationOptions };

const NAME_TIMING_FN = qname('p', 'timing', NS.pml);
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

// Largest numeric `<p:cTn id="N">` anywhere in the tree (0 when none).
const maxCTnId = (el: XmlElement): number => {
  let max = 0;
  const walk = (e: XmlElement): void => {
    if (isPml(e, 'cTn')) {
      const n = Number.parseInt(getAttrValue(e, ATTR_ID_FN) ?? '', 10);
      if (Number.isFinite(n) && n > max) max = n;
    }
    for (const c of e.children) if (c.kind === 'element') walk(c);
  };
  walk(el);
  return max;
};

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

// Merges a freshly-built single-effect timing into an existing `<p:timing>`,
// renumbering the new effect's cTn ids so they stay unique. Returns false when
// the existing tree lacks the mainSeq structure we know how to extend (so the
// caller can avoid destroying it). This is what lets a second shape animate
// without wiping a template's pre-existing animations.
const mergeEffectInto = (existing: XmlElement, fresh: XmlElement): boolean => {
  const existingMainSeqChildTnLst = (() => {
    const mainSeq = findDescendant(
      existing,
      (e) => isPml(e, 'cTn') && getAttrValue(e, qname('', 'nodeType', '')) === 'mainSeq',
    );
    return mainSeq ? firstChildElement(mainSeq, qname('p', 'childTnLst', NS.pml)) : null;
  })();
  if (!existingMainSeqChildTnLst) return false;

  // The fresh tree's click-effect wrapper is the <p:par> under its own mainSeq
  // childTnLst. Lift it out and renumber its cTn ids past the existing max.
  const freshMainSeq = findDescendant(
    fresh,
    (e) => isPml(e, 'cTn') && getAttrValue(e, qname('', 'nodeType', '')) === 'mainSeq',
  );
  const freshChildTnLst = freshMainSeq
    ? firstChildElement(freshMainSeq, qname('p', 'childTnLst', NS.pml))
    : null;
  const newPar = freshChildTnLst
    ? freshChildTnLst.children.find((c): c is XmlElement => c.kind === 'element' && isPml(c, 'par'))
    : null;
  const freshBldP = findDescendant(fresh, (e) => isPml(e, 'bldP'));
  if (!newPar || !freshBldP) return false;

  const offset = maxCTnId(existing) - 2; // fresh effect ids start at 3
  if (offset > 0) shiftCTnIds(newPar, offset);

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
  if (effectCTn) setGrpId(effectCTn, newGrpId);
  setGrpId(freshBldP, newGrpId);

  existingMainSeqChildTnLst.children.push(newPar);
  const existingBldLst = findDescendant(existing, (e) => isPml(e, 'bldLst'));
  if (existingBldLst) existingBldLst.children.push(freshBldP);
  else existing.children.push(elem(qname('p', 'bldLst', NS.pml), { children: [freshBldP] }));
  return true;
};

/**
 * Sets a single click-triggered animation effect on the given shape.
 *
 * The effect is *merged* into any existing `<p:timing>` on the slide rather
 * than replacing it: animating a second shape (or re-running on a template that
 * already has authored animations) preserves the existing effects and appends
 * this one as the next click stop, with cTn ids renumbered to stay unique. To
 * clear every animation first, call `clearSlideAnimations`.
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
 */
export const setShapeAnimation = (shape: SlideShapeData, opts: AnimationOptions): void => {
  const slide = shape[SHAPE_SLIDE];
  const spid = shape[SHAPE_SNAPSHOT].id;
  const fresh = buildSingleEffectTiming(spid, opts);
  const existing = findTiming(slide);
  if (existing === null) {
    insertTimingAtEnd(slide, fresh);
  } else if (!mergeEffectInto(existing, fresh)) {
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
 * `<p:timing>` tree, or `null` if the shape has no animation in the
 * v1 single-effect schema we model. Unknown presets are reported as a
 * raw `null` rather than guessing.
 */
export const getShapeAnimation = (shape: SlideShapeData): AnimationEffect | null => {
  const slide = shape[SHAPE_SLIDE];
  const timing = slide[SLIDE_DOCUMENT].root.children.find(
    (c): c is XmlElement =>
      c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'timing',
  );
  if (!timing) return null;

  // Confirm the shape's spid appears in <p:bldLst><p:bldP spid="..."/>.
  const bldLst = firstChildElement(timing, qname('p', 'bldLst', NS.pml));
  if (!bldLst) return null;
  const spidStr = String(shape[SHAPE_SNAPSHOT].id);
  const matched = allChildElements(bldLst, qname('p', 'bldP', NS.pml)).some(
    (b) => getAttrValue(b, qname('', 'spid', '')) === spidStr,
  );
  if (!matched) return null;

  // Walk the timing tree to find the effect cTn for this shape. Our
  // builder emits `<p:cTn presetID="N" presetClass="entr|exit" ...
  // nodeType="clickEffect">` with a `<p:spTgt spid="..."/>` inside. We
  // accept any cTn carrying that combination.
  let presetID: string | null = null;
  let presetClass: string | null = null;
  const walk = (el: XmlElement): boolean => {
    if (el.name.namespaceURI === NS.pml && el.name.localName === 'cTn') {
      const cls = getAttrValue(el, qname('', 'presetClass', ''));
      const id = getAttrValue(el, qname('', 'presetID', ''));
      if (cls && id) {
        // Confirm this cTn targets our shape via a descendant spTgt.
        const targetsShape = (sub: XmlElement): boolean => {
          if (
            sub.name.namespaceURI === NS.pml &&
            sub.name.localName === 'spTgt' &&
            getAttrValue(sub, qname('', 'spid', '')) === spidStr
          ) {
            return true;
          }
          for (const c of sub.children) {
            if (c.kind === 'element' && targetsShape(c)) return true;
          }
          return false;
        };
        if (targetsShape(el)) {
          presetClass = cls;
          presetID = id;
          return true;
        }
      }
    }
    for (const c of el.children) {
      if (c.kind === 'element' && walk(c)) return true;
    }
    return false;
  };
  walk(timing);
  if (!presetID || !presetClass) return null;

  // Map back to AnimationEffect.
  const id = Number.parseInt(presetID, 10);
  if (presetClass === 'entr' && id === 1) return 'appear';
  if (presetClass === 'entr' && id === 10) return 'fadeIn';
  if (presetClass === 'exit' && id === 1) return 'disappear';
  if (presetClass === 'exit' && id === 10) return 'fadeOut';
  return null;
};

/** Removes the slide's `<p:timing>` element entirely. */
/**
 * Returns every shape on the slide that has an authored animation
 * effect (i.e. `getShapeAnimation(shape)` is not `null`). Pair to
 * `slideHasAnimations`. Useful for audit reports — "which shapes on
 * this slide actually animate?" before exporting to a video pipeline
 * that doesn't honor PowerPoint's timing tree.
 */
export const findShapesWithAnimation = (slide: SlideData): ReadonlyArray<SlideShapeData> => {
  const out: SlideShapeData[] = [];
  for (const shape of slide[SLIDE_SHAPES]) {
    if (getShapeAnimation(shape) !== null) out.push(shape);
  }
  return out;
};

export const clearSlideAnimations = (slide: SlideData): void => {
  removeExistingTiming(slide);
  commitSlideData(slide);
  refreshSlideData(slide);
};

void NAME_TIMING_FN;
