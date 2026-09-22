// Editing the animation effects a slide already has.
//
// Every effect is addressed by the `<p:cTn id>` `getSlideAnimations` reports,
// which survives these calls: reordering and retiming rewrite the click stops
// and groups around an effect, never the effect's own id. A change is assembled
// in a copy of `<p:timing>` and swapped in once, so a refusal leaves the slide
// exactly as it was.

import { unsignedIntMs } from '../../internal/bounds.ts';
import {
  type AnimationEffect,
  type AnimationOptions,
  type AnimationStartCondition,
  buildSingleEffectTiming,
  isMediaTimingNode,
} from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
  cloneElement,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';
import { SLIDE_DOCUMENT, type SlideData } from '../_internal-symbols.ts';
import {
  type CTnIds,
  type LayoutStep,
  type MainSeqLayout,
  buildKeyOf,
  cTnIdsUnder,
  dropShapesFromTiming,
  idsAfter,
  isEmptyTiming,
  isPlainEffect,
  largestCTnId,
  pruneBuildEntries,
  readMainSeqLayout,
  referencedNodeIds,
  targetedSpids,
  setBuildByParagraph,
  shiftCTnIds,
  writeMainSeqLayout,
} from './_animation-layout.ts';
import {
  type AnimationStepNode,
  findSlideTimingElement,
  readTimingSteps,
  setEffectDelayMs,
  setEffectDurationMs,
} from './_animation-timing.ts';
import { commitSlideData, refreshSlideData } from './_helpers.ts';
import { rootChildTnLst } from './_media-timing.ts';

/**
 * What to change about an effect. Every field is optional; the ones left out
 * keep what the slide already says.
 *
 * `byParagraph` is not an attribute flip: turning it on replaces the effect
 * with one per paragraph of the shape's text, and turning it off replaces the
 * whole build with a single effect on the shape. In both directions the id you
 * passed stays on the effect that takes the original's place in the click
 * order, and the other effects of a build are new — so a caller holding several
 * handles into one build keeps only the one it addressed.
 */
export interface AnimationPatch {
  readonly effect?: AnimationEffect;
  readonly durationMs?: number;
  readonly start?: AnimationStartCondition;
  readonly delayMs?: number;
  readonly byParagraph?: boolean;
}

const ATTR_GRP_ID = qname('', 'grpId', '');

const NO_TIMING = 'the slide has no animation effects';
const NOT_EDITABLE =
  'is one this library cannot edit safely — it uses a preset, target or start condition that is ' +
  'not modelled here. Read it with getSlideAnimations and leave it as it is';
const NOT_LAYOUTABLE =
  "the slide's main animation sequence is nested in a way this library cannot lay out again, so " +
  'reordering it could change what the other effects do';

/**
 * Runs `edit` against a copy of the slide's timing and installs the result.
 * `edit` throws to refuse, which leaves the slide untouched — the copy is what
 * makes an edit that touches several effects all-or-nothing.
 */
const editTiming = (slide: SlideData, fn: string, edit: (timing: XmlElement) => void): void => {
  const existing = findSlideTimingElement(slide);
  if (existing === null) throw new Error(`${fn}: ${NO_TIMING}.`);
  const draft = cloneElement(existing);
  edit(draft);
  const children = slide[SLIDE_DOCUMENT].root.children;
  const index = children.indexOf(existing);
  // Removing the last effect leaves an empty `<p:timing>`, which is not what a
  // slide without animations looks like.
  if (isEmptyTiming(draft)) children.splice(index, 1);
  else children[index] = draft;
  commitSlideData(slide);
  refreshSlideData(slide);
};

interface Located {
  readonly layout: MainSeqLayout;
  readonly index: number;
  readonly node: AnimationStepNode;
  /** The `<p:par>` each step lives in, so a build's paragraphs are reachable without re-scanning. */
  readonly parByCTn: ReadonlyMap<XmlElement, XmlElement>;
}

const locate = (timing: XmlElement, id: number, fn: string): Located => {
  const node = readTimingSteps(timing).find((n) => n.step.id === id);
  if (node === undefined) throw new Error(`${fn}: the slide has no animation with id ${id}.`);
  if (!node.step.editable) throw new Error(`${fn}: animation ${id} ${NOT_EDITABLE}.`);
  const layout = readMainSeqLayout(timing);
  const index = layout === null ? -1 : layout.steps.findIndex((s) => s.cTn === node.cTn);
  if (layout === null || index < 0) throw new Error(`${fn}: ${NOT_LAYOUTABLE}.`);
  return {
    layout,
    index,
    node,
    parByCTn: new Map(layout.steps.map((s) => [s.cTn, s.par])),
  };
};

/**
 * Refuses to drop time nodes another part of the tree points at. An effect's
 * own id and its behaviours' ids can be the target of a `<p:tn val>` condition
 * elsewhere, and removing them would leave that condition waiting on nothing.
 */
const refuseIfReferenced = (
  timing: XmlElement,
  going: ReadonlySet<string>,
  keep: ReadonlySet<string>,
  what: string,
  fn: string,
): void => {
  const referenced = referencedNodeIds(timing);
  for (const id of going) {
    if (!keep.has(id) && referenced.has(id)) {
      throw new Error(
        `${fn}: another time node on this slide starts from ${what} (<p:tn val="${id}">), so ` +
          'removing it would leave that condition waiting on nothing.',
      );
    }
  }
};

// The `<p:sp>` whose non-visual properties carry this id. Walks the whole
// shape tree, so a shape inside a group is found too.
const findShapeElement = (el: XmlElement, id: string): XmlElement | null => {
  for (const child of el.children) {
    if (child.kind !== 'element') continue;
    const nv = firstChildElement(child, qname('p', 'nvSpPr', NS.pml));
    const cNvPr = nv === null ? null : firstChildElement(nv, qname('p', 'cNvPr', NS.pml));
    if (cNvPr !== null && getAttrValue(cNvPr, qname('', 'id', '')) === id) return child;
    const found = findShapeElement(child, id);
    if (found !== null) return found;
  }
  return null;
};

/** Number of paragraphs in the text body of the shape with this id. */
const paragraphCount = (slide: SlideData, spid: number): number => {
  const shape = findShapeElement(slide[SLIDE_DOCUMENT].root, String(spid));
  const txBody = shape === null ? null : firstChildElement(shape, qname('p', 'txBody', NS.pml));
  if (txBody === null) return 0;
  return txBody.children.filter(
    (c) => c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'p',
  ).length;
};

/** The paragraph a step targets, or `null` when it animates the whole shape. */
const paragraphOf = (node: AnimationStepNode, fn: string): number | null => {
  const target = node.step.target;
  if (target.kind !== 'paragraphs') return null;
  if (target.firstParagraph !== target.lastParagraph) {
    throw new Error(
      `${fn}: animation ${node.step.id} covers paragraphs ${target.firstParagraph}-` +
        `${target.lastParagraph} at once, which this library cannot rewrite one paragraph at a ` +
        'time. Remove it and add the effects you want instead.',
    );
  }
  return target.firstParagraph;
};

/**
 * Builds the replacement effect for `node`, keeping its id and build group.
 * Behaviour ids are taken from past the end of the tree so they cannot collide
 * with anything already in it.
 */
const rebuiltStep = (
  node: AnimationStepNode,
  opts: AnimationOptions,
  paragraph: number | null,
  keepId: number | null,
  fn: string,
  ids: CTnIds,
): LayoutStep => {
  const spid = node.step.target.shapeId;
  const fresh = buildSingleEffectTiming(spid!, opts, { paragraph, label: fn });
  const step = readMainSeqLayout(fresh)!.steps[0]!;
  // The builder numbers a standalone effect 5. None of these effects is in the
  // tree yet, so they take their ids from the edit's own counter rather than
  // from the tree — otherwise every paragraph of a build would be numbered the
  // same.
  shiftCTnIds(step.par, ids.next - 5);
  ids.next = largestCTnId(step.par) + 1;
  if (keepId !== null) setAttr(step.cTn, 'id', String(keepId));
  const grpId = getAttrValue(node.cTn, ATTR_GRP_ID);
  if (grpId !== null) setAttr(step.cTn, 'grpId', grpId);
  return step;
};

const setAttr = (el: XmlElement, name: string, value: string): void => {
  el.attrs = el.attrs.map((a) =>
    a.name.namespaceURI === '' && a.name.localName === name ? { ...a, value } : a,
  );
};

/** The steps of the build `node` belongs to, in click order — itself when it is not in one. */
const buildSiblings = (
  timing: XmlElement,
  layout: MainSeqLayout,
  node: AnimationStepNode,
): AnimationStepNode[] => {
  const key = buildKeyOf(node.cTn, node.step.target.shapeId!);
  const inLayout = new Set(layout.steps.map((s) => s.cTn));
  return readTimingSteps(timing).filter(
    (n) =>
      inLayout.has(n.cTn) &&
      n.step.target.kind === 'paragraphs' &&
      n.step.target.shapeId === node.step.target.shapeId &&
      buildKeyOf(n.cTn, n.step.target.shapeId) === key,
  );
};

/**
 * Changes one effect: its preset, how long it runs, when it starts, and whether
 * the shape's text is revealed a paragraph at a time.
 *
 * The effect keeps the id it was addressed by. Everything after it in the click
 * order is laid out again, because a longer effect moves what follows it and a
 * new start condition moves the effect itself between click stops.
 *
 * An effect whose duration the slide never stated gets this library's default
 * when it is rewritten; pass `durationMs` to say what it should be.
 */
export const updateSlideAnimation = (slide: SlideData, id: number, patch: AnimationPatch): void => {
  const fn = 'updateSlideAnimation';
  editTiming(slide, fn, (timing) => {
    const { layout, index, node, parByCTn } = locate(timing, id, fn);
    const current = layout.steps[index]!;
    const paragraph = paragraphOf(node, fn);
    const wasBuild = paragraph !== null;
    const nowBuild = patch.byParagraph ?? wasBuild;
    const start = patch.start ?? current.start;
    const spid = node.step.target.shapeId!;
    const ids = idsAfter(timing);

    // Changing the preset or the paragraph build means writing a new effect
    // node; anything else is set on the one that is there, so an effect this
    // library did not author keeps whatever else it carries.
    const changesPreset = patch.effect !== undefined && patch.effect !== node.step.effect;
    const rebuild = changesPreset || nowBuild !== wasBuild;

    let replaced: ReadonlySet<XmlElement> = new Set([node.cTn]);
    let replacements: LayoutStep[];

    if (!rebuild) {
      if (
        patch.delayMs !== undefined &&
        !setEffectDelayMs(node.cTn, boundedMs(patch.delayMs, `${fn}: delayMs`))
      ) {
        throw new Error(`${fn}: animation ${id} ${NOT_PLAIN_START}.`);
      }
      if (
        patch.durationMs !== undefined &&
        !setEffectDurationMs(node.cTn, boundedMs(patch.durationMs, `${fn}: durationMs`))
      ) {
        throw new Error(`${fn}: animation ${id} ${NOT_ONE_DURATION}.`);
      }
      replacements = [{ ...current, start }];
    } else {
      if (!isPlainEffect(current.par)) {
        throw new Error(`${fn}: animation ${id} ${NOT_REPLACEABLE}.`);
      }
      // The effect node keeps its id, but its behaviours are written anew.
      refuseIfReferenced(
        timing,
        cTnIdsUnder(current.par),
        new Set([String(id)]),
        `a behaviour of animation ${id}`,
        fn,
      );
      const durationMs = patch.durationMs ?? node.step.durationMs;
      const delayMs = patch.delayMs ?? node.step.delayMs;
      const opts: AnimationOptions = {
        effect: patch.effect ?? node.step.effect!,
        ...(durationMs === null ? {} : { durationMs }),
        ...(delayMs === null ? {} : { delayMs }),
        start,
      };
      if (nowBuild === wasBuild) {
        replacements = [rebuiltStep(node, opts, paragraph, id, fn, ids)];
      } else if (nowBuild) {
        const count = paragraphCount(slide, spid);
        if (count === 0) {
          throw new Error(
            `${fn}: byParagraph needs a shape with text. Shape ${spid} has no paragraphs to build.`,
          );
        }
        // The addressed effect becomes the first paragraph and keeps its id;
        // the rest are new effects, so only that one handle stays valid.
        replacements = Array.from({ length: count }, (_, i) =>
          rebuiltStep(node, opts, i, i === 0 ? id : null, fn, ids),
        );
      } else {
        // Collapsing a build: every paragraph of it goes, and the single effect
        // that replaces them takes the addressed step's place in the order.
        const siblings = buildSiblings(timing, layout, node);
        const pars = siblings.map((n) => parByCTn.get(n.cTn)!);
        if (!pars.every(isPlainEffect)) {
          throw new Error(`${fn}: animation ${id} ${NOT_REPLACEABLE}.`);
        }
        const going = new Set(pars.flatMap((par) => [...cTnIdsUnder(par)]));
        refuseIfReferenced(timing, going, new Set([String(id)]), `a paragraph of this build`, fn);
        replaced = new Set(siblings.map((n) => n.cTn));
        replacements = [rebuiltStep(node, opts, null, id, fn, ids)];
      }
    }

    const steps = layout.steps.flatMap((step) =>
      step.cTn === node.cTn ? replacements : replaced.has(step.cTn) ? [] : [step],
    );
    if (!writeMainSeqLayout(timing, layout, steps, ids)) throw new Error(`${fn}: ${CANNOT_CHAIN}`);
    if (nowBuild !== wasBuild) setBuildByParagraph(timing, buildKeyOf(node.cTn, spid), nowBuild);
  });
};

// The builder is the one place that bounds a time value; reuse it so a patch
// and a fresh effect refuse the same inputs.
const boundedMs = (value: number, label: string): number => unsignedIntMs(value, label);

const NOT_PLAIN_START =
  'does not start at a plain offset — its start condition ties it to something else, so a delay ' +
  'would not mean what you asked for';
const NOT_ONE_DURATION =
  'animates through more than one timed behaviour, so it has no single length to set. Change the ' +
  'effect instead, or leave its timing as it is';
const NOT_REPLACEABLE =
  'carries timing this library did not author, and changing its preset or its paragraph build ' +
  'would replace the effect and lose that. Remove it and add the effect you want instead';

const CANNOT_CHAIN =
  'an effect set to start after the one before it now follows one whose length this library ' +
  'cannot measure, so there is no moment to start it at. Give that effect a duration, or start ' +
  'this one on click.';

/**
 * Removes one effect. The effects after it move one place earlier in the click
 * order and are retimed to match.
 *
 * A paragraph of a build is removed on its own: the build entry the rest of the
 * paragraphs share stays, and goes only once the last of them is gone.
 */
export const removeSlideAnimation = (slide: SlideData, id: number): void => {
  const fn = 'removeSlideAnimation';
  editTiming(slide, fn, (timing) => {
    const { layout, node, parByCTn } = locate(timing, id, fn);
    const key = buildKeyOf(node.cTn, node.step.target.shapeId!);
    refuseIfReferenced(
      timing,
      cTnIdsUnder(parByCTn.get(node.cTn)!),
      new Set(),
      `animation ${id}`,
      fn,
    );
    const steps = layout.steps.filter((s) => s.cTn !== node.cTn);
    if (!writeMainSeqLayout(timing, layout, steps, idsAfter(timing)))
      throw new Error(`${fn}: ${CANNOT_CHAIN}`);
    pruneBuildEntries(timing, new Set([key]));
  });
};

/**
 * Moves one effect to `index` in the click order, counting only the slide's
 * main sequence — the order `getSlideAnimations` reports for steps whose
 * `sequence` is `'mainSeq'`.
 *
 * Each effect keeps the start condition it had, so moving a `'withPrevious'`
 * effect to the front does not turn it into a click: it becomes the first thing
 * that runs as the slide appears, which is what it already meant.
 */
export const moveSlideAnimation = (slide: SlideData, id: number, index: number): void => {
  const fn = 'moveSlideAnimation';
  editTiming(slide, fn, (timing) => {
    const located = locate(timing, id, fn);
    const { layout } = located;
    if (!Number.isInteger(index) || index < 0 || index >= layout.steps.length) {
      throw new Error(
        `${fn}: index ${index} is outside the slide's ${layout.steps.length} main-sequence ` +
          'animations.',
      );
    }
    const steps = [...layout.steps];
    const [moved] = steps.splice(located.index, 1);
    steps.splice(index, 0, moved!);
    if (!writeMainSeqLayout(timing, layout, steps, idsAfter(timing)))
      throw new Error(`${fn}: ${CANNOT_CHAIN}`);
  });
};

/**
 * Works out what the slide's timing has to become once the shapes in `spids`
 * are gone, and returns the step that installs it. Throws instead when the
 * timing could not be left consistent, which is why it runs before the shapes
 * themselves are taken out: the caller has changed nothing yet, so a refusal
 * costs the slide nothing.
 *
 * Nothing cascades. An effect on a shape that is staying can be timed against
 * one that is going, and quietly taking that effect away too would delete an
 * animation the caller never asked about. The deletion is refused instead.
 *
 * @internal
 */
export const planShapeAnimationRemoval = (
  slide: SlideData,
  spids: ReadonlySet<number>,
): (() => void) => {
  const existing = findSlideTimingElement(slide);
  if (existing === null || spids.size === 0) return NO_ANIMATION_REMOVAL;
  const draft = cloneElement(existing);
  const steps = readTimingSteps(draft);
  const shared = steps.find(
    (n) =>
      n.step.targetShapeIds.some((id) => spids.has(id)) &&
      n.step.targetShapeIds.some((id) => !spids.has(id)),
  );
  if (shared !== undefined) {
    throw new Error(
      `${REMOVE_SHAPE}: one animation on this slide drives both this shape and a shape that is ` +
        'staying, and this library cannot take one of them out of it. Remove that animation ' +
        'first with removeSlideAnimation.',
    );
  }

  const doomed = new Set(
    steps.filter((n) => n.step.targetShapeIds.some((id) => spids.has(id))).map((n) => n.cTn),
  );
  const before = cTnIdsUnder(draft);
  // The checks below run even when nothing is dropped: a shape can be named by
  // the timing without any effect animating it — as the trigger of an
  // interactive sequence, say — and that reference dangles just the same.
  const changed = dropShapesFromTiming(draft, spids, doomed, REMOVE_SHAPE);

  const after = cTnIdsUnder(draft);
  const dangling = [...referencedNodeIds(draft)].find((id) => !after.has(id) && before.has(id));
  if (dangling !== undefined) {
    throw new Error(
      `${REMOVE_SHAPE}: another animation on this slide is timed against one of this shape's ` +
        `effects (<p:tn val="${dangling}">). Remove that animation first, or the file would be ` +
        'left waiting on a time node that is gone.',
    );
  }

  // Anything still naming a removed shape is a reference this library does not
  // model — the trigger of an interactive sequence, say — and leaving it would
  // point the file at a shape that is not there. Media time nodes are not read:
  // they are the play controls of a clip, and the caller takes them off itself.
  const named = animatedSpids(draft);
  const stranded = [...spids].find((spid) => named.has(String(spid)));
  if (stranded !== undefined) {
    throw new Error(
      `${REMOVE_SHAPE}: the slide's animation timing refers to shape ${stranded} in a way this ` +
        'library does not model, so removing the shape would leave that reference dangling. ' +
        "Clear the slide's animations first with clearSlideAnimations.",
    );
  }
  if (!changed) return NO_ANIMATION_REMOVAL;

  return () => {
    const children = slide[SLIDE_DOCUMENT].root.children;
    const index = children.indexOf(existing);
    if (index < 0) return;
    if (isEmptyTiming(draft)) children.splice(index, 1);
    else children[index] = draft;
  };
};

const NO_ANIMATION_REMOVAL = (): void => {};

const REMOVE_SHAPE = 'removeShape';

/** Every shape the timing's sequences name, play controls left out of it. */
const animatedSpids = (timing: XmlElement): Set<string> => {
  const rootList = rootChildTnLst(timing);
  const out = new Set<string>();
  if (rootList === null) return out;
  for (const child of rootList.children) {
    if (child.kind !== 'element' || isMediaTimingNode(child)) continue;
    for (const spid of targetedSpids(child)) out.add(spid);
  }
  return out;
};
