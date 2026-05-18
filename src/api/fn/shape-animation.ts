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

const removeExistingTiming = (slide: SlideData): void => {
  slide[SLIDE_DOCUMENT].root.children = slide[SLIDE_DOCUMENT].root.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'timing'),
  );
};

const insertTimingAtEnd = (slide: SlideData, timing: XmlElement): void => {
  // Schema ordering: `<p:timing>` is one of the last children of `<p:sld>`
  // (after cSld, clrMapOvr, transition). Appending to the end of
  // `<p:sld>` keeps the file valid.
  slide[SLIDE_DOCUMENT].root.children.push(timing);
};

/**
 * Sets a single click-triggered animation effect on the given shape.
 * Replaces any existing `<p:timing>` block on the slide — v1 supports
 * exactly one effect per slide. Calling this on a second shape replaces
 * the first.
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
  removeExistingTiming(slide);
  const spid = shape[SHAPE_SNAPSHOT].id;
  const timing = buildSingleEffectTiming(spid, opts);
  insertTimingAtEnd(slide, timing);
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
