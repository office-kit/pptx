// Slide animations.

import { unsignedIntMs } from '../../internal/bounds.ts';

import {
  type AnimationEffect,
  type AnimationOptions,
  buildSingleEffectTiming,
  buildTimingRoot,
} from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
  allChildElements,
  attr,
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
import { maxCTnId, mediaTimingNodes, rootChildTnLst } from './_media-timing.ts';
// ---------------------------------------------------------------------------
// Animation authoring appends click-triggered entrance/exit effects.
// Read/edit helpers preserve separate effects and unrelated timing data.

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

// Merges a freshly-built single-effect timing into an existing `<p:timing>`,
// renumbering the new effect's cTn ids so they stay unique. Returns false when
// the existing tree has no structure we know how to extend (so the caller can
// avoid destroying it). This is what lets a second shape animate without wiping
// a template's pre-existing animations.
const mergeEffectInto = (existing: XmlElement, fresh: XmlElement): boolean => {
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
  let existingMainSeqChildTnLst = existingMainSeq
    ? firstChildElement(existingMainSeq, qname('p', 'childTnLst', NS.pml))
    : null;
  if (existingMainSeq && !existingMainSeqChildTnLst) {
    existingMainSeqChildTnLst = elem(qname('p', 'childTnLst', NS.pml));
    const sub = firstChildElement(existingMainSeq, qname('p', 'subTnLst', NS.pml));
    if (sub)
      existingMainSeq.children.splice(
        existingMainSeq.children.indexOf(sub),
        0,
        existingMainSeqChildTnLst,
      );
    else existingMainSeq.children.push(existingMainSeqChildTnLst);
  }
  if (existingMainSeqChildTnLst) {
    // The fresh tree's click-effect wrapper is the <p:par> under its own mainSeq
    // childTnLst. Lift it out and renumber its cTn ids past the existing max.
    const offset = maxCTnId(existing) - 2; // fresh effect ids start at 3
    if (offset > 0) shiftCTnIds(newPar, offset);
    existingMainSeqChildTnLst.children.push(newPar);
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

/** Moves an effect one position, retaining its start mode and behavior subtree. */
export const moveSlideAnimation = (
  slide: SlideData,
  timingId: string,
  direction: 'earlier' | 'later',
): void => {
  if (direction !== 'earlier' && direction !== 'later')
    throw new Error('Invalid animation move direction.');
  const timing = findTiming(slide);
  const main = timing && findDescendant(timing, isMainSeqCTn);
  const list = main && firstChildElement(main, qname('p', 'childTnLst', NS.pml));
  const groups = list?.children.filter((node): node is XmlElement => node.kind === 'element') ?? [];
  const effects = groups.map((group) => {
    const found: XmlElement[] = [];
    const visit = (node: XmlElement): void => {
      if (isPml(node, 'cTn') && getAttrValue(node, qname('', 'presetID', '')) !== null)
        found.push(node);
      for (const child of node.children) if (child.kind === 'element') visit(child);
    };
    visit(group);
    return found;
  });
  const index = effects.findIndex((group) =>
    group.some((effect) => getAttrValue(effect, ATTR_ID_FN) === timingId),
  );
  if (index < 0 || !list || !timing) throw new Error('Animation effect no longer exists.');
  if (
    effects.some(
      (group) =>
        group.length !== 1 || getAttrValue(group[0]!, qname('', 'nodeType', '')) !== 'clickEffect',
    )
  ) {
    const copy = structuredClone(timing);
    const sequence = readSimpleAnimationSequence(copy);
    const from = sequence.effects.findIndex(
      (effect) => getAttrValue(effect, ATTR_ID_FN) === timingId,
    );
    const to = from + (direction === 'earlier' ? -1 : 1);
    if (to < 0 || to >= sequence.effects.length) return;
    const owners = new Map<string, XmlElement>();
    const collect = (node: XmlElement, owner: XmlElement): void => {
      if (isPml(node, 'cTn')) {
        const id = getAttrValue(node, ATTR_ID_FN);
        if (id !== null) owners.set(id, owner);
      }
      for (const child of node.children) if (child.kind === 'element') collect(child, owner);
    };
    for (const effect of sequence.effects) collect(effect, effect);
    const moved = new Set([sequence.effects[from]!, sequence.effects[to]!]);
    const verify = (node: XmlElement, owner?: XmlElement): void => {
      if (sequence.effects.includes(node)) owner = node;
      if (isPml(node, 'tn')) {
        const target = owners.get(getAttrValue(node, qname('', 'val', '')) ?? '');
        if (target !== owner && ((target && moved.has(target)) || (owner && moved.has(owner))))
          throw new Error('Other animation timing depends on the order of these effects.');
      }
      for (const child of node.children) if (child.kind === 'element') verify(child, owner);
    };
    verify(copy);
    [sequence.effects[from], sequence.effects[to]] = [
      sequence.effects[to]!,
      sequence.effects[from]!,
    ];
    rebuildSimpleAnimationSequence(copy, sequence.list, sequence.effects);
    const root = slide[SLIDE_DOCUMENT].root;
    root.children[root.children.indexOf(timing)] = copy;
    commitSlideData(slide);
    refreshSlideData(slide);
    return;
  }
  const destination = index + (direction === 'earlier' ? -1 : 1);
  if (destination < 0 || destination >= groups.length) return;
  for (const position of [index, destination]) {
    const group = effects[position]!;
    if (group.length !== 1 || getAttrValue(group[0]!, qname('', 'nodeType', '')) !== 'clickEffect')
      throw new Error('Reordering grouped animation effects is not supported yet.');
    let container = groups[position]!;
    while (container !== group[0]) {
      let child: XmlElement | null;
      if (isPml(container, 'par') || isPml(container, 'childTnLst')) {
        const children = container.children.filter(
          (node): node is XmlElement => node.kind === 'element',
        );
        child = children.length === 1 ? children[0]! : null;
      } else if (
        isPml(container, 'cTn') &&
        !firstChildElement(container, qname('p', 'subTnLst', NS.pml))
      ) {
        child = firstChildElement(container, qname('p', 'childTnLst', NS.pml));
      } else child = null;
      if (!child)
        throw new Error('Reordering this animation timing structure is not supported yet.');
      container = child;
    }
  }
  // References across a moving subtree boundary can change meaning when reordered.
  // Validate all references before mutating the existing tree.
  const membership = new Map<string, number>();
  const collect = (node: XmlElement, group: number): void => {
    if (isPml(node, 'cTn')) {
      const id = getAttrValue(node, ATTR_ID_FN);
      if (id !== null) membership.set(id, group);
    }
    for (const child of node.children) if (child.kind === 'element') collect(child, group);
  };
  collect(groups[index]!, index);
  collect(groups[destination]!, destination);
  const verify = (node: XmlElement, owner = -1): void => {
    if (node === groups[index]) owner = index;
    if (node === groups[destination]) owner = destination;
    if (isPml(node, 'tn')) {
      const target = membership.get(getAttrValue(node, qname('', 'val', '')) ?? '') ?? -1;
      if (target !== owner && (target !== -1 || owner !== -1))
        throw new Error('Other animation timing depends on the order of these effects.');
    }
    for (const child of node.children) if (child.kind === 'element') verify(child, owner);
  };
  verify(timing);
  const from = list.children.indexOf(groups[index]!);
  const to = list.children.indexOf(groups[destination]!);
  [list.children[from], list.children[to]] = [list.children[to]!, list.children[from]!];
  commitSlideData(slide);
  refreshSlideData(slide);
};

/** Removes one main-sequence effect and its empty wrappers, retaining unrelated timing. */
export const removeSlideAnimation = (slide: SlideData, timingId: string): void => {
  const original = findTiming(slide);
  if (!original) throw new Error('Animation effect no longer exists.');
  const timing = structuredClone(original);
  const main = findDescendant(timing, isMainSeqCTn);
  const list = main && firstChildElement(main, qname('p', 'childTnLst', NS.pml));
  const parents = new Map<XmlElement, XmlElement>();
  const visit = (node: XmlElement): void => {
    for (const child of node.children)
      if (child.kind === 'element') {
        parents.set(child, node);
        visit(child);
      }
  };
  visit(timing);
  const effect =
    list &&
    findDescendant(
      list,
      (node) =>
        isPml(node, 'cTn') &&
        getAttrValue(node, ATTR_ID_FN) === timingId &&
        getAttrValue(node, qname('', 'presetID', '')) !== null,
    );
  if (!effect) throw new Error('Animation effect no longer exists.');
  if (
    effect.children.some(
      (child) =>
        child.kind === 'element' &&
        findDescendant(
          child,
          (node) => isPml(node, 'cTn') && getAttrValue(node, qname('', 'presetID', '')) !== null,
        ),
    )
  )
    throw new Error('Nested animation effects cannot be removed independently.');
  const elements = (node: XmlElement) =>
    node.children.filter((child): child is XmlElement => child.kind === 'element');
  let removed = parents.get(effect)!;
  if (!isPml(removed, 'par') || elements(removed).length !== 1)
    throw new Error('Unsupported animation effect container.');
  // Prune only empty timing wrappers on the selected path, stopping at mainSeq.
  for (;;) {
    const ownerList = parents.get(removed);
    if (
      !ownerList ||
      ownerList === list ||
      !isPml(ownerList, 'childTnLst') ||
      elements(ownerList).length !== 1
    )
      break;
    const ownerTime = parents.get(ownerList);
    const ownerPar = ownerTime && parents.get(ownerTime);
    if (
      !ownerTime ||
      !isPml(ownerTime, 'cTn') ||
      getAttrValue(ownerTime, qname('', 'presetID', '')) !== null ||
      firstChildElement(ownerTime, qname('p', 'subTnLst', NS.pml)) ||
      !ownerPar ||
      !isPml(ownerPar, 'par') ||
      elements(ownerPar).length !== 1
    )
      break;
    removed = ownerPar;
  }
  const removedIds = new Set<string>();
  const collect = (node: XmlElement): void => {
    if (isPml(node, 'cTn')) {
      const id = getAttrValue(node, ATTR_ID_FN);
      if (id !== null) removedIds.add(id);
    }
    for (const child of elements(node)) collect(child);
  };
  collect(removed);
  let simple: ReturnType<typeof readSimpleAnimationSequence> | undefined;
  try {
    simple = readSimpleAnimationSequence(timing);
  } catch (cause) {
    if (!(cause instanceof SimpleAnimationSequenceError)) throw cause;
  }
  if (simple) {
    // Only the selected effect's own nodes disappear. Rebuilt wrappers reconnect
    // later After Previous effects to their new preceding parallel group.
    removedIds.clear();
    collect(effect);
    rebuildSimpleAnimationSequence(
      timing,
      simple.list,
      simple.effects.filter((node) => node !== effect),
    );
  }
  const owner = parents.get(removed)!;
  if (!simple) owner.children = owner.children.filter((child) => child !== removed);
  if (!simple && !elements(owner).length) {
    const parent = parents.get(owner)!;
    parent.children = parent.children.filter((child) => child !== owner);
  }
  if (
    findDescendant(
      timing,
      (node) => isPml(node, 'tn') && removedIds.has(getAttrValue(node, qname('', 'val', '')) ?? ''),
    )
  )
    throw new Error('Other animation timing depends on this effect.');
  const group = getAttrValue(effect, qname('', 'grpId', ''));
  if (
    group !== null &&
    !findDescendant(
      timing,
      (node) => isPml(node, 'cTn') && getAttrValue(node, qname('', 'grpId', '')) === group,
    )
  ) {
    const builds = firstChildElement(timing, qname('p', 'bldLst', NS.pml));
    if (builds) {
      builds.children = builds.children.filter(
        (child) =>
          child.kind !== 'element' || getAttrValue(child, qname('', 'grpId', '')) !== group,
      );
      if (!elements(builds).length)
        timing.children = timing.children.filter((child) => child !== builds);
    }
  }
  const root = slide[SLIDE_DOCUMENT].root;
  root.children[root.children.indexOf(original)] = timing;
  commitSlideData(slide);
  refreshSlideData(slide);
};

export type AnimationStart = 'clickEffect' | 'withEffect' | 'afterEffect';

export interface AnimationSettings {
  effect?: AnimationEffect;
  start?: AnimationStart;
  durationMs?: number;
  delayMs?: number;
}

/** Applies settings to a selection, restoring all timing on any failure. */
export const setSlideAnimationsSettings = (
  slide: SlideData,
  timingIds: readonly string[],
  settings: AnimationSettings,
): void => {
  const ids = [...new Set(timingIds)];
  if (!ids.length) return;
  const existing = new Set(
    getSlideAnimationSequence(slide)
      .flat()
      .map((effect) => effect.timingId),
  );
  if (ids.some((id) => !existing.has(id))) throw new Error('Animation effect no longer exists.');
  const original = structuredClone(findTiming(slide)!);
  try {
    for (const id of ids) {
      if (settings.effect !== undefined) setSlideAnimationEffect(slide, id, settings.effect);
      if (settings.start !== undefined) setSlideAnimationStart(slide, id, settings.start);
      if (settings.durationMs !== undefined)
        setSlideAnimationDuration(slide, id, settings.durationMs);
      if (settings.delayMs !== undefined) setSlideAnimationDelay(slide, id, settings.delayMs);
    }
  } catch (cause) {
    const root = slide[SLIDE_DOCUMENT].root;
    root.children[root.children.indexOf(findTiming(slide)!)] = original;
    commitSlideData(slide);
    refreshSlideData(slide);
    throw cause;
  }
};

/** Moves selected blocks by one position, preserving their internal order. */
export const moveSlideAnimations = (
  slide: SlideData,
  timingIds: readonly string[],
  direction: 'earlier' | 'later',
): void => {
  if (direction !== 'earlier' && direction !== 'later')
    throw new Error('Invalid animation move direction.');
  const selected = new Set(timingIds);
  if (!selected.size) return;
  const order = getSlideAnimationSequence(slide)
    .flat()
    .map((effect) => effect.timingId!);
  if ([...selected].some((id) => !order.includes(id)))
    throw new Error('Animation effect no longer exists.');
  const original = structuredClone(findTiming(slide)!);
  const step = direction === 'earlier' ? -1 : 1;
  const traversal = direction === 'earlier' ? [...order] : [...order].reverse();
  try {
    for (const id of traversal) {
      if (!selected.has(id)) continue;
      const index = order.indexOf(id),
        neighbor = index + step;
      if (neighbor < 0 || neighbor >= order.length || selected.has(order[neighbor]!)) continue;
      moveSlideAnimation(slide, id, direction);
      [order[index], order[neighbor]] = [order[neighbor]!, order[index]!];
    }
  } catch (cause) {
    const root = slide[SLIDE_DOCUMENT].root;
    root.children[root.children.indexOf(findTiming(slide)!)] = original;
    commitSlideData(slide);
    refreshSlideData(slide);
    throw cause;
  }
};

/** Places selected effects before a target effect, or at the end for null. */
export const reorderSlideAnimations = (
  slide: SlideData,
  timingIds: readonly string[],
  beforeId: string | null,
): void => {
  const selected = new Set(timingIds);
  const order = getSlideAnimationSequence(slide)
    .flat()
    .map((effect) => effect.timingId!);
  if (
    [...selected].some((id) => !order.includes(id)) ||
    (beforeId !== null && !order.includes(beforeId))
  )
    throw new Error('Animation effect no longer exists.');
  if (!selected.size || (beforeId !== null && selected.has(beforeId))) return;
  const remaining = order.filter((id) => !selected.has(id));
  const insertion = beforeId === null ? remaining.length : remaining.indexOf(beforeId);
  const desired = [
    ...remaining.slice(0, insertion),
    ...order.filter((id) => selected.has(id)),
    ...remaining.slice(insertion),
  ];
  const original = structuredClone(findTiming(slide)!);
  try {
    for (let index = 0; index < desired.length; index++) {
      let from = order.indexOf(desired[index]!);
      while (from > index) {
        moveSlideAnimation(slide, desired[index]!, 'earlier');
        [order[from - 1], order[from]] = [order[from]!, order[from - 1]!];
        from--;
      }
    }
  } catch (cause) {
    const root = slide[SLIDE_DOCUMENT].root;
    root.children[root.children.indexOf(findTiming(slide)!)] = original;
    commitSlideData(slide);
    refreshSlideData(slide);
    throw cause;
  }
};

/** Removes a selection as one atomic operation. */
export const removeSlideAnimations = (slide: SlideData, timingIds: readonly string[]): void => {
  const ids = [...new Set(timingIds)];
  if (!ids.length) return;
  const existing = new Set(
    getSlideAnimationSequence(slide)
      .flat()
      .map((effect) => effect.timingId),
  );
  if (ids.some((id) => !existing.has(id))) throw new Error('Animation effect no longer exists.');
  const original = findTiming(slide)!;
  try {
    for (const id of ids) removeSlideAnimation(slide, id);
  } catch (cause) {
    const root = slide[SLIDE_DOCUMENT].root;
    const current = findTiming(slide);
    if (current) root.children[root.children.indexOf(current)] = original;
    else root.children.push(original);
    commitSlideData(slide);
    refreshSlideData(slide);
    throw cause;
  }
};

/** Replaces an effect's behaviors while retaining its identity and start conditions. */
export const setSlideAnimationEffect = (
  slide: SlideData,
  timingId: string,
  effect: AnimationEffect,
): void => {
  const original = findTiming(slide);
  const metadata = getSlideAnimationSequence(slide)
    .flat()
    .find((item) => item.timingId === timingId);
  if (!original || !metadata) throw new Error('Animation effect no longer exists.');
  const timing = structuredClone(original);
  const selected = findDescendant(
    timing,
    (node) => isPml(node, 'cTn') && getAttrValue(node, ATTR_ID_FN) === timingId,
  )!;
  const children = firstChildElement(selected, qname('p', 'childTnLst', NS.pml));
  if (!children) throw new Error('Animation effect has no replaceable behaviors.');
  const targets: XmlElement[] = [];
  const removedIds = new Set<string>();
  const collect = (node: XmlElement): void => {
    if (isPml(node, 'spTgt')) targets.push(node);
    if (isPml(node, 'cTn')) {
      const id = getAttrValue(node, ATTR_ID_FN);
      if (id !== null) removedIds.add(id);
      if (getAttrValue(node, qname('', 'presetID', '')) !== null)
        throw new Error('Nested animation effects cannot be replaced independently.');
    }
    for (const child of node.children) if (child.kind === 'element') collect(child);
  };
  collect(children);
  const target = targets[0];
  if (!target || targets.some((item) => JSON.stringify(item) !== JSON.stringify(target)))
    throw new Error('Effect replacement requires a single consistent object target.');
  const spid = Number(getAttrValue(target, qname('', 'spid', '')));
  if (!Number.isInteger(spid) || spid < 1) throw new Error('Invalid animation object target.');
  const verify = (node: XmlElement): void => {
    if (node === children) return;
    if (isPml(node, 'tn') && removedIds.has(getAttrValue(node, qname('', 'val', '')) ?? ''))
      throw new Error('Other animation timing depends on these effect behaviors.');
    for (const child of node.children) if (child.kind === 'element') verify(child);
  };
  verify(timing);
  const fresh = buildSingleEffectTiming(spid, {
    effect,
    durationMs:
      metadata.effect === 'fadeIn' || metadata.effect === 'fadeOut'
        ? (metadata.durationMs ?? 500)
        : 500,
  });
  shiftCTnIds(fresh, maxCTnId(timing));
  const replacement = findDescendant(
    fresh,
    (node) => isPml(node, 'cTn') && getAttrValue(node, qname('', 'presetID', '')) !== null,
  )!;
  const behaviors = firstChildElement(replacement, qname('p', 'childTnLst', NS.pml))!;
  const copyTarget = (node: XmlElement): void => {
    node.children = node.children.map((child) => {
      if (child.kind !== 'element') return child;
      if (isPml(child, 'spTgt')) return structuredClone(target);
      copyTarget(child);
      return child;
    });
  };
  copyTarget(behaviors);
  const presetAttributes = ['presetID', 'presetClass', 'presetSubtype'];
  selected.attrs = selected.attrs.filter(
    (item) => item.name.namespaceURI !== '' || !presetAttributes.includes(item.name.localName),
  );
  selected.attrs.push(
    ...replacement.attrs.filter(
      (item) => item.name.namespaceURI === '' && presetAttributes.includes(item.name.localName),
    ),
  );
  selected.children[selected.children.indexOf(children)] = behaviors;
  const root = slide[SLIDE_DOCUMENT].root;
  root.children[root.children.indexOf(original)] = timing;
  commitSlideData(slide);
  refreshSlideData(slide);
};

class SimpleAnimationSequenceError extends Error {}

// Validates before rebuilding so unsupported native timing remains untouched.
const readSimpleAnimationSequence = (timing: XmlElement) => {
  const main = findDescendant(timing, isMainSeqCTn);
  const list = main && firstChildElement(main, qname('p', 'childTnLst', NS.pml));
  if (!list) throw new SimpleAnimationSequenceError('Animation effect no longer exists.');
  const effects: XmlElement[] = [];
  const wrappers = new Set<XmlElement>();
  const wrapperIds = new Set<string>();
  const parents = new Map<XmlElement, XmlElement>();
  const value = (node: XmlElement, name: string) => getAttrValue(node, qname('', name, ''));
  const elements = (node: XmlElement) =>
    node.children.filter((n): n is XmlElement => n.kind === 'element');
  const unsupported = () => {
    throw new SimpleAnimationSequenceError('Editing requires a simple main animation sequence.');
  };
  const inspect = (node: XmlElement): void => {
    if (isPml(node, 'cTn') && value(node, 'presetID') !== null) {
      if (!['clickEffect', 'withEffect', 'afterEffect'].includes(value(node, 'nodeType') ?? ''))
        unsupported();
      if (
        node.children.some(
          (child) =>
            child.kind === 'element' &&
            findDescendant(child, (n) => isPml(n, 'cTn') && value(n, 'presetID') !== null),
        )
      )
        unsupported();
      effects.push(node);
      return;
    }
    if (!['par', 'cTn', 'childTnLst', 'stCondLst', 'cond', 'tn'].some((name) => isPml(node, name)))
      unsupported();
    wrappers.add(node);
    const allowed = isPml(node, 'cTn')
      ? ['id', 'fill']
      : isPml(node, 'cond')
        ? ['delay', 'evt']
        : isPml(node, 'tn')
          ? ['val']
          : [];
    if (node.attrs.some((a) => a.name.namespaceURI !== '' || !allowed.includes(a.name.localName)))
      unsupported();
    if (isPml(node, 'cTn')) {
      if (value(node, 'fill') !== null && value(node, 'fill') !== 'hold') unsupported();
      if (value(node, 'id') !== null) wrapperIds.add(value(node, 'id')!);
    }
    if (
      isPml(node, 'cond') &&
      (!['0', 'indefinite', null].includes(value(node, 'delay')) ||
        ![null, 'onEnd'].includes(value(node, 'evt')))
    )
      unsupported();
    if (isPml(node, 'stCondLst') && elements(node).length !== 1) unsupported();
    for (const child of elements(node)) {
      parents.set(child, node);
      inspect(child);
    }
  };
  for (const node of elements(list)) {
    const count = effects.length;
    inspect(node);
    // An empty stop can carry navigation/timing meaning and is not ours to drop.
    if (effects.length === count) unsupported();
  }
  // Wrapper IDs will be regenerated. References owned by effects or other sequences
  // must never be silently detached from their original timing nodes.
  const checkReferences = (node: XmlElement): void => {
    if (
      isPml(node, 'tn') &&
      (wrappers.has(node)
        ? !wrapperIds.has(value(node, 'val') ?? '')
        : wrapperIds.has(value(node, 'val') ?? ''))
    )
      unsupported();
    for (const child of elements(node)) checkReferences(child);
  };
  checkReferences(timing);
  for (const node of wrappers) {
    if (!isPml(node, 'tn')) continue;
    const cond = parents.get(node),
      conditions = cond && parents.get(cond);
    const owner = conditions && parents.get(conditions),
      par = owner && parents.get(owner);
    const siblings = par && parents.get(par);
    const preceding =
      siblings && par ? elements(siblings)[elements(siblings).indexOf(par) - 1] : undefined;
    const precedingTime = preceding && firstChildElement(preceding, qname('p', 'cTn', NS.pml));
    if (
      !cond ||
      value(cond, 'evt') !== 'onEnd' ||
      !precedingTime ||
      value(precedingTime, 'id') !== value(node, 'val')
    )
      unsupported();
  }
  return { list, effects };
};

const rebuildSimpleAnimationSequence = (
  timing: XmlElement,
  list: XmlElement,
  effects: XmlElement[],
): void => {
  const value = (node: XmlElement, name: string) => getAttrValue(node, qname('', name, ''));
  let nextId = maxCTnId(timing) + 1;
  const condition = (delay: string, previous?: string) =>
    elem(qname('p', 'stCondLst', NS.pml), {
      children: [
        elem(qname('p', 'cond', NS.pml), {
          attrs: [
            attr(qname('', 'delay', ''), delay),
            ...(previous ? [attr(qname('', 'evt', ''), 'onEnd')] : []),
          ],
          children: previous
            ? [elem(qname('p', 'tn', NS.pml), { attrs: [attr(qname('', 'val', ''), previous)] })]
            : [],
        }),
      ],
    });
  const wrapper = (delay: string, previous?: string) => {
    const children = elem(qname('p', 'childTnLst', NS.pml));
    const id = String(nextId++);
    const par = elem(qname('p', 'par', NS.pml), {
      children: [
        elem(qname('p', 'cTn', NS.pml), {
          attrs: [attr(ATTR_ID_FN, id), attr(qname('', 'fill', ''), 'hold')],
          children: [condition(delay, previous), children],
        }),
      ],
    });
    return { par, children, id };
  };
  const groups: XmlElement[] = [];
  let group: ReturnType<typeof wrapper> | undefined;
  let cohort: ReturnType<typeof wrapper> | undefined;
  for (const effect of effects) {
    const trigger = value(effect, 'nodeType');
    if (!group || trigger === 'clickEffect') {
      group = wrapper(trigger === 'clickEffect' ? 'indefinite' : '0');
      groups.push(group.par);
      cohort = undefined;
    }
    if (!cohort || trigger === 'afterEffect') {
      cohort = wrapper('0', cohort?.id);
      group.children.children.push(cohort.par);
    }
    cohort.children.children.push(elem(qname('p', 'par', NS.pml), { children: [effect] }));
  }
  list.children = groups;
};

/** Changes a main-sequence start mode while retaining effect and build subtrees. */
export const setSlideAnimationStart = (
  slide: SlideData,
  timingId: string,
  start: AnimationStart,
): void => {
  if (!['clickEffect', 'withEffect', 'afterEffect'].includes(start))
    throw new Error('Invalid animation start mode.');
  const original = findTiming(slide);
  if (!original) throw new Error('Animation effect no longer exists.');
  const timing = structuredClone(original);
  const { list, effects } = readSimpleAnimationSequence(timing);
  const selected = effects.find((node) => getAttrValue(node, ATTR_ID_FN) === timingId);
  if (!selected) throw new Error('Animation effect no longer exists.');
  selected.attrs = selected.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'nodeType'),
  );
  selected.attrs.push(attr(qname('', 'nodeType', ''), start));
  rebuildSimpleAnimationSequence(timing, list, effects);
  const root = slide[SLIDE_DOCUMENT].root;
  root.children[root.children.indexOf(original)] = timing;
  commitSlideData(slide);
  refreshSlideData(slide);
};

/** Updates a simple effect start delay without changing its trigger or behaviors. */
export const setSlideAnimationDelay = (
  slide: SlideData,
  timingId: string,
  delayMs: number,
): void => {
  const delay = unsignedIntMs(delayMs, 'setSlideAnimationDelay: delayMs');
  const timing = findTiming(slide);
  const main = timing && findDescendant(timing, isMainSeqCTn);
  const effect =
    main &&
    findDescendant(
      main,
      (node) =>
        isPml(node, 'cTn') &&
        getAttrValue(node, ATTR_ID_FN) === timingId &&
        getAttrValue(node, qname('', 'presetID', '')) !== null,
    );
  if (!effect) throw new Error('Animation effect no longer exists.');
  let conditions = firstChildElement(effect, qname('p', 'stCondLst', NS.pml));
  const items =
    conditions?.children.filter((node): node is XmlElement => node.kind === 'element') ?? [];
  const condition = items[0];
  if (
    conditions &&
    (items.length !== 1 ||
      !condition ||
      !isPml(condition, 'cond') ||
      condition.children.some((node) => node.kind === 'element') ||
      getAttrValue(condition, qname('', 'evt', '')) !== null ||
      !/^\d*$/.test(getAttrValue(condition, qname('', 'delay', '')) ?? ''))
  )
    throw new Error('Delay editing requires a simple numeric start condition.');
  if (condition) {
    condition.attrs = condition.attrs.filter(
      (item) => !(item.name.namespaceURI === '' && item.name.localName === 'delay'),
    );
    condition.attrs.push(attr(qname('', 'delay', ''), String(delay)));
  } else {
    conditions = elem(qname('p', 'stCondLst', NS.pml), {
      children: [
        elem(qname('p', 'cond', NS.pml), { attrs: [attr(qname('', 'delay', ''), String(delay))] }),
      ],
    });
    effect.children.unshift(conditions);
  }
  commitSlideData(slide);
  refreshSlideData(slide);
};

// PowerPoint uses animEffect/filter="fade"; our authoring also supports anim/style.opacity.
const fadeBehaviorTime = (node: XmlElement, presetClass: string | null): XmlElement | null => {
  if (!isPml(node, 'anim') && !isPml(node, 'animEffect')) return null;
  const behavior = firstChildElement(node, qname('p', 'cBhvr', NS.pml));
  if (!behavior) return null;
  if (isPml(node, 'animEffect')) {
    const direction = presetClass === 'entr' ? 'in' : presetClass === 'exit' ? 'out' : null;
    if (
      !direction ||
      getAttrValue(node, qname('', 'filter', '')) !== 'fade' ||
      getAttrValue(node, qname('', 'transition', '')) !== direction ||
      firstChildElement(node, qname('p', 'progress', NS.pml))
    )
      return null;
  } else {
    const names = firstChildElement(behavior, qname('p', 'attrNameLst', NS.pml));
    if (
      !names ||
      !findDescendant(
        names,
        (item) =>
          isPml(item, 'attrName') &&
          item.children.some((child) => child.kind === 'text' && child.data === 'style.opacity'),
      )
    )
      return null;
  }
  return firstChildElement(behavior, qname('p', 'cTn', NS.pml));
};

/**
 * Updates a supported fade in place, preserving its targets, ordering,
 * timing IDs and other effects. Unsupported timing behaviors are left intact.
 */
export const setSlideAnimationDuration = (
  slide: SlideData,
  timingId: string,
  durationMs: number,
): void => {
  const duration = unsignedIntMs(durationMs, 'setSlideAnimationDuration: durationMs');
  const timing = findTiming(slide);
  const effect =
    timing &&
    findDescendant(
      timing,
      (node) =>
        isPml(node, 'cTn') &&
        getAttrValue(node, ATTR_ID_FN) === timingId &&
        getAttrValue(node, qname('', 'presetID', '')) !== null,
    );
  if (!effect) throw new Error('Animation effect no longer exists.');
  const presetClass = getAttrValue(effect, qname('', 'presetClass', ''));
  if (
    getAttrValue(effect, qname('', 'presetID', '')) !== '10' ||
    !['entr', 'exit'].includes(presetClass ?? '')
  )
    throw new Error('Duration editing is not supported for this animation effect.');
  const opacityTimes: XmlElement[] = [];
  const hiddenConditions: XmlElement[] = [];
  const walk = (node: XmlElement): void => {
    if (
      node !== effect &&
      isPml(node, 'cTn') &&
      getAttrValue(node, qname('', 'presetID', '')) !== null
    )
      return;
    const fadeTime = fadeBehaviorTime(node, presetClass);
    if (fadeTime) opacityTimes.push(fadeTime);
    if (isPml(node, 'set')) {
      const behavior = firstChildElement(node, qname('p', 'cBhvr', NS.pml));
      const time = behavior && firstChildElement(behavior, qname('p', 'cTn', NS.pml));
      const names = behavior && firstChildElement(behavior, qname('p', 'attrNameLst', NS.pml));
      const hasName = (name: string): boolean =>
        !!names &&
        !!findDescendant(
          names,
          (item) =>
            isPml(item, 'attrName') &&
            item.children.some((child) => child.kind === 'text' && child.data === name),
        );
      if (
        time &&
        isPml(node, 'set') &&
        hasName('style.visibility') &&
        findDescendant(
          node,
          (item) => isPml(item, 'strVal') && getAttrValue(item, qname('', 'val', '')) === 'hidden',
        )
      ) {
        const conditions = firstChildElement(time, qname('p', 'stCondLst', NS.pml));
        if (conditions)
          hiddenConditions.push(...allChildElements(conditions, qname('p', 'cond', NS.pml)));
      }
    }
    for (const child of node.children) if (child.kind === 'element') walk(child);
  };
  walk(effect);
  if (opacityTimes.length !== 1)
    throw new Error('Duration editing requires a single supported fade animation.');
  const time = opacityTimes[0]!;
  const oldDuration = getAttrValue(time, qname('', 'dur', ''));
  if (oldDuration === null || !/^\d+$/.test(oldDuration))
    throw new Error('Animation duration is not numeric.');
  const setValue = (node: XmlElement, name: string, value: string): void => {
    node.attrs = node.attrs.filter(
      (item) => !(item.name.namespaceURI === '' && item.name.localName === name),
    );
    node.attrs.push(attr(qname('', name, ''), value));
  };
  setValue(time, 'dur', String(duration));
  // Our opacity exits carry a visibility switch at the fade endpoint.
  for (const condition of hiddenConditions)
    if (getAttrValue(condition, qname('', 'delay', '')) === oldDuration)
      setValue(condition, 'delay', String(duration));
  commitSlideData(slide);
  refreshSlideData(slide);
};

/** A preset effect in the main slide sequence, in document execution order. */
export interface SlideAnimationEffect {
  readonly timingId: string | null;
  readonly shapeIds: readonly string[];
  readonly presetId: string | null;
  readonly presetClass: string | null;
  /** Unknown presets remain present with a null effect. */
  readonly effect: AnimationEffect | null;
  readonly trigger: string | null;
  /** Authored effect-node delay; null means absent or non-numeric. */
  readonly delayMs: number | null;
  /** Fade behavior duration; zero for instant supported effects. */
  readonly durationMs: number | null;
}

/**
 * Reads main-sequence groups in timing-tree order, independently of shape z-order.
 * Each group is a direct mainSeq child (normally one click stop); effects inside
 * retain their clickEffect/withEffect/afterEffect trigger. Interactive sequences
 * are excluded. This is timing metadata, not a general timing-tree evaluator.
 */
export const getSlideAnimationSequence = (
  slide: SlideData,
): ReadonlyArray<ReadonlyArray<SlideAnimationEffect>> => {
  const timing = findTiming(slide);
  const main = timing && findDescendant(timing, isMainSeqCTn);
  const list = main && firstChildElement(main, qname('p', 'childTnLst', NS.pml));
  if (!list) return [];
  const attribute = (node: XmlElement, name: string): string | null =>
    getAttrValue(node, qname('', name, ''));
  const milliseconds = (value: string | null): number | null =>
    value !== null && /^\d+$/.test(value) && Number(value) <= 4294967295 ? Number(value) : null;
  return list.children
    .filter((node): node is XmlElement => node.kind === 'element')
    .map((group) => {
      const effects: SlideAnimationEffect[] = [];
      const visit = (node: XmlElement): void => {
        if (isPml(node, 'cTn') && attribute(node, 'presetID') !== null) {
          const presetId = attribute(node, 'presetID');
          const presetClass = attribute(node, 'presetClass');
          const effect: AnimationEffect | null =
            presetClass === 'entr'
              ? presetId === '1'
                ? 'appear'
                : presetId === '10'
                  ? 'fadeIn'
                  : null
              : presetClass === 'exit'
                ? presetId === '1'
                  ? 'disappear'
                  : presetId === '10'
                    ? 'fadeOut'
                    : null
                : null;
          const shapeIds = new Set<string>();
          const fadeTimes: XmlElement[] = [];
          const inspect = (child: XmlElement): void => {
            // Nested effect nodes own their targets and behaviors.
            if (child !== node && isPml(child, 'cTn') && attribute(child, 'presetID') !== null)
              return;
            if (isPml(child, 'spTgt')) {
              const id = attribute(child, 'spid');
              if (id !== null) shapeIds.add(id);
            }
            const time = fadeBehaviorTime(child, presetClass);
            if (time) fadeTimes.push(time);
            for (const item of child.children) if (item.kind === 'element') inspect(item);
          };
          inspect(node);
          const conditions = firstChildElement(node, qname('p', 'stCondLst', NS.pml));
          const condition = conditions && firstChildElement(conditions, qname('p', 'cond', NS.pml));
          effects.push({
            timingId: attribute(node, 'id'),
            shapeIds: [...shapeIds],
            presetId,
            presetClass,
            effect,
            trigger: attribute(node, 'nodeType'),
            delayMs: condition ? milliseconds(attribute(condition, 'delay')) : null,
            durationMs:
              effect === 'appear' || effect === 'disappear'
                ? 0
                : fadeTimes.length === 1
                  ? milliseconds(attribute(fadeTimes[0]!, 'dur'))
                  : null,
          });
        }
        for (const child of node.children) if (child.kind === 'element') visit(child);
      };
      visit(group);
      return effects;
    });
};

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

void NAME_TIMING_FN;
