// Laying out a slide's main animation sequence.
//
// Adding an effect appends to whatever structure a slide already has, but
// moving, retiming or removing one cannot be done in place: the click stops and
// the groups inside them are what encode "on click", "with previous" and "after
// previous", so changing any of those means deciding where every stop and group
// falls all over again. This module reads that structure out as a flat list of
// steps and writes a new one back from a list in the order the caller wants.
//
// Only `<p:seq nodeType="mainSeq">` is touched. Interactive sequences, media
// time nodes and anything else under the timing root are left exactly as they
// were, and each step's `<p:par>` moves whole — nothing inside an effect this
// library does not model is rewritten.

import {
  NS,
  type XmlElement,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';
import type { AnimationStartCondition } from '../../internal/presentationml/index.ts';
import { groupEndMs } from './_animation-timing.ts';
import { rootChildTnLst } from './_media-timing.ts';

const NAME_PAR = qname('p', 'par', NS.pml);
const NAME_C_TN = qname('p', 'cTn', NS.pml);
const NAME_CHILD_TN_LST = qname('p', 'childTnLst', NS.pml);
const NAME_ST_COND_LST = qname('p', 'stCondLst', NS.pml);
const NAME_COND = qname('p', 'cond', NS.pml);
const NAME_BLD_LST = qname('p', 'bldLst', NS.pml);
const NAME_TN = qname('p', 'tn', NS.pml);

const ATTR_ID = qname('', 'id', '');
const ATTR_FILL = qname('', 'fill', '');
const ATTR_DELAY = qname('', 'delay', '');
const ATTR_NODE_TYPE = qname('', 'nodeType', '');
const ATTR_GRP_ID = qname('', 'grpId', '');
const ATTR_SPID = qname('', 'spid', '');
const ATTR_VAL = qname('', 'val', '');

const isPml = (el: XmlElement, local: string): boolean =>
  el.name.namespaceURI === NS.pml && el.name.localName === local;

// ST_TLTimeNodeType tokens, both ways. The node type is what PowerPoint's UI
// reads back as the effect's start, so a layout that moves a step has to keep
// the two in step.
const START_OF_NODE_TYPE: Record<string, AnimationStartCondition> = {
  clickEffect: 'click',
  withEffect: 'withPrevious',
  afterEffect: 'afterPrevious',
};
const NODE_TYPE_OF_START: Record<AnimationStartCondition, string> = {
  click: 'clickEffect',
  withPrevious: 'withEffect',
  afterPrevious: 'afterEffect',
};

/** One effect in the main sequence, with the start it runs on. */
export interface LayoutStep {
  readonly par: XmlElement;
  readonly cTn: XmlElement;
  readonly start: AnimationStartCondition;
}

/** The main sequence as a flat list of steps in click order. */
export interface MainSeqLayout {
  readonly childTnLst: XmlElement;
  readonly steps: readonly LayoutStep[];
}

const elementChildren = (el: XmlElement | null): XmlElement[] =>
  el === null ? [] : el.children.filter((c): c is XmlElement => c.kind === 'element');

/**
 * Everything this library puts on a click stop or a group. A wrapper carrying
 * anything else means something the layout does not model — it repeats, ends on
 * a condition, iterates, or hangs an extension off itself — and rewriting it
 * from scratch would drop that silently.
 */
const WRAPPER_ATTRS = new Set(['id', 'fill']);

const isPlainWrapper = (par: XmlElement): boolean => {
  if (par.attrs.length > 0 || elementChildren(par).length !== 1) return false;
  const cTn = firstChildElement(par, NAME_C_TN);
  if (cTn === null) return false;
  if (!cTn.attrs.every((a) => a.name.namespaceURI === '' && WRAPPER_ATTRS.has(a.name.localName))) {
    return false;
  }
  // The wrapper is re-emitted with `fill="hold"`; any other fill would be
  // changed by that, so it is not a wrapper we can reproduce.
  if (getAttrValue(cTn, ATTR_FILL) !== 'hold') return false;
  if (!elementChildren(cTn).every((c) => isPml(c, 'stCondLst') || isPml(c, 'childTnLst'))) {
    return false;
  }
  const stCondLst = firstChildElement(cTn, NAME_ST_COND_LST);
  if (stCondLst === null) return false;
  const conds = elementChildren(stCondLst);
  const cond = conds[0];
  if (conds.length !== 1 || cond === undefined || !isPml(cond, 'cond')) return false;
  // A condition with a child points at another node's lifetime, and one with
  // `evt` fires on an event; neither survives being re-emitted as an offset.
  if (elementChildren(cond).length > 0) return false;
  return cond.attrs.every((a) => a.name.namespaceURI === '' && a.name.localName === 'delay');
};

/** Time node ids something else in the tree points at with `<p:tn val>`. */
export const referencedNodeIds = (timing: XmlElement): Set<string> => {
  const ids = new Set<string>();
  const walk = (el: XmlElement): void => {
    if (el.name.namespaceURI === NAME_TN.namespaceURI && el.name.localName === 'tn') {
      const val = getAttrValue(el, ATTR_VAL);
      if (val !== null) ids.add(val);
    }
    for (const c of el.children) if (c.kind === 'element') walk(c);
  };
  walk(timing);
  return ids;
};

/** Every `<p:cTn id>` at or under `el` — the ids that go when `el` goes. */
export const cTnIdsUnder = (el: XmlElement): Set<string> => {
  const ids = new Set<string>();
  const walk = (e: XmlElement): void => {
    if (isPml(e, 'cTn')) {
      const id = getAttrValue(e, ATTR_ID);
      if (id !== null) ids.add(id);
    }
    for (const c of e.children) if (c.kind === 'element') walk(c);
  };
  walk(el);
  return ids;
};

/** The `<p:par>` children of a wrapper's `<p:childTnLst>`, or `null` for anything else in it. */
const parChildren = (par: XmlElement): XmlElement[] | null => {
  const cTn = firstChildElement(par, NAME_C_TN);
  const childTnLst = cTn === null ? null : firstChildElement(cTn, NAME_CHILD_TN_LST);
  const children = elementChildren(childTnLst);
  return children.every((c) => isPml(c, 'par')) ? children : null;
};

const mainSeqCTn = (timing: XmlElement): XmlElement | null => {
  let found: XmlElement | null = null;
  const walk = (el: XmlElement): void => {
    if (found !== null) return;
    if (isPml(el, 'cTn') && getAttrValue(el, ATTR_NODE_TYPE) === 'mainSeq') {
      found = el;
      return;
    }
    for (const c of el.children) if (c.kind === 'element') walk(c);
  };
  walk(timing);
  return found;
};

/**
 * Reads the main sequence as click stop > group > effect, the shape this
 * library writes and PowerPoint emits.
 *
 * `null` when the sequence is nested some other way, or holds an effect whose
 * `nodeType` does not say when it starts. Laying such a tree out again would
 * have to invent a start for it, so the caller refuses the edit instead.
 */
export const readMainSeqLayout = (timing: XmlElement): MainSeqLayout | null => {
  const seq = mainSeqCTn(timing);
  if (seq === null) return null;
  const childTnLst = firstChildElement(seq, NAME_CHILD_TN_LST);
  if (childTnLst === null) return { childTnLst: emptyChildTnLst(seq), steps: [] };

  const stops = elementChildren(childTnLst);
  if (!stops.every((c) => isPml(c, 'par'))) return null;
  // The wrappers are replaced wholesale when the sequence is laid out again, so
  // they may hold nothing we would lose and nothing may point at their ids.
  const referenced = referencedNodeIds(timing);
  const replaceable = (par: XmlElement): boolean =>
    isPlainWrapper(par) &&
    !referenced.has(getAttrValue(firstChildElement(par, NAME_C_TN)!, ATTR_ID) ?? '');

  const steps: LayoutStep[] = [];
  for (const [stopIndex, stop] of stops.entries()) {
    if (!replaceable(stop)) return null;
    const groups = parChildren(stop);
    if (groups === null || groups.length === 0) return null;

    const perGroup: LayoutStep[][] = [];
    for (const group of groups) {
      if (!replaceable(group)) return null;
      const effects = parChildren(group);
      if (effects === null || effects.length === 0) return null;
      const groupSteps: LayoutStep[] = [];
      for (const par of effects) {
        const cTn = firstChildElement(par, NAME_C_TN);
        if (cTn === null) return null;
        const start = START_OF_NODE_TYPE[getAttrValue(cTn, ATTR_NODE_TYPE) ?? ''];
        if (start === undefined) return null;
        groupSteps.push({ par, cTn, start });
      }
      perGroup.push(groupSteps);
    }

    // Laying the sequence out again rewrites every offset, so the ones already
    // there have to be the ones we would write. A stop that waits 250ms, or a
    // group placed somewhere other than the end of the one before it, means
    // something this library did not author and cannot reproduce.
    const opening = perGroup[0]![0]!.start;
    if (stopIndex > 0 && opening !== 'click') return null;
    if (delayOf(stop) !== (opening === 'click' ? 'indefinite' : '0')) return null;
    for (const [groupIndex, group] of groups.entries()) {
      const expected = groupIndex === 0 ? '0' : stringifyEnd(groupEndMs(groups[groupIndex - 1]!));
      if (expected === null || delayOf(group) !== expected) return null;
      // Only the group that opens a stop may hold a step that is not "after"
      // the one before it; everything else joins the group it runs with.
      if (groupIndex > 0 && perGroup[groupIndex]![0]!.start !== 'afterPrevious') return null;
      steps.push(...perGroup[groupIndex]!);
    }
  }
  return { childTnLst, steps };
};

const delayOf = (par: XmlElement): string | null => {
  const cTn = firstChildElement(par, NAME_C_TN);
  const stCondLst = cTn === null ? null : firstChildElement(cTn, NAME_ST_COND_LST);
  const cond = stCondLst === null ? null : firstChildElement(stCondLst, NAME_COND);
  return cond === null ? null : getAttrValue(cond, ATTR_DELAY);
};

const stringifyEnd = (end: number | null): string | null => (end === null ? null : String(end));

// CT_TimeNodeList needs a child, so a sequence with no effects has no
// `<p:childTnLst>` at all. One is added the moment a step goes in.
const emptyChildTnLst = (seq: XmlElement): XmlElement => {
  const childTnLst = elem(NAME_CHILD_TN_LST);
  seq.children.push(childTnLst);
  return childTnLst;
};

export const largestCTnId = (el: XmlElement): number => {
  let max = 0;
  const walk = (e: XmlElement): void => {
    if (isPml(e, 'cTn')) {
      const raw = getAttrValue(e, ATTR_ID);
      if (raw !== null && /^\d+$/.test(raw)) max = Math.max(max, Number(raw));
    }
    for (const c of e.children) if (c.kind === 'element') walk(c);
  };
  walk(el);
  return max;
};

/**
 * Hands out `<p:cTn id>` values for one edit. Every node a call adds — the
 * wrappers it lays out and the behaviours of any effect it rebuilds — takes its
 * id from the same counter, because none of them is in the tree yet when the
 * next one is numbered.
 */
export interface CTnIds {
  next: number;
}

export const idsAfter = (timing: XmlElement): CTnIds => ({ next: largestCTnId(timing) + 1 });

/** A click stop or group: a `<p:par>` whose only job is to start its children at `delay`. */
const wrapperPar = (id: number, delay: string): XmlElement =>
  elem(NAME_PAR, {
    children: [
      elem(NAME_C_TN, {
        attrs: [attr(ATTR_ID, String(id)), attr(ATTR_FILL, 'hold')],
        children: [
          elem(NAME_ST_COND_LST, {
            children: [elem(NAME_COND, { attrs: [attr(ATTR_DELAY, delay)] })],
          }),
          elem(NAME_CHILD_TN_LST),
        ],
      }),
    ],
  });

const appendInto = (wrapper: XmlElement, child: XmlElement): void => {
  const cTn = firstChildElement(wrapper, NAME_C_TN)!;
  firstChildElement(cTn, NAME_CHILD_TN_LST)!.children.push(child);
};

const setNodeType = (cTn: XmlElement, start: AnimationStartCondition): void => {
  const value = NODE_TYPE_OF_START[start];
  const has = cTn.attrs.some((a) => a.name.namespaceURI === '' && a.name.localName === 'nodeType');
  cTn.attrs = has
    ? cTn.attrs.map((a) =>
        a.name.namespaceURI === '' && a.name.localName === 'nodeType' ? { ...a, value } : a,
      )
    : [...cTn.attrs, attr(ATTR_NODE_TYPE, value)];
};

/**
 * Writes `steps` back as the main sequence, in the order given.
 *
 * A `'click'` step opens a stop that waits for the viewer; a stop opened by a
 * leading `'withPrevious'` / `'afterPrevious'` step starts at zero instead, so
 * a sequence that begins with one still runs as the slide appears rather than
 * being quietly turned into a click. `'withPrevious'` joins the group before
 * it and `'afterPrevious'` opens the next group in the same stop, offset to
 * where that group ends — the same clock the writing path uses.
 *
 * `false` when an `'afterPrevious'` step follows one whose length cannot be
 * measured. The caller is working on a copy, so nothing is written.
 */
export const writeMainSeqLayout = (
  timing: XmlElement,
  layout: MainSeqLayout,
  steps: readonly LayoutStep[],
  ids: CTnIds,
): boolean => {
  // Effect ids stay as they are — they are the handles callers hold — so the
  // wrappers take ids from past everything the edit has used so far.
  const stops: XmlElement[] = [];
  let stop: XmlElement | null = null;
  let group: XmlElement | null = null;

  for (const step of steps) {
    if (stop === null || step.start === 'click') {
      stop = wrapperPar(ids.next++, step.start === 'click' ? 'indefinite' : '0');
      group = wrapperPar(ids.next++, '0');
      appendInto(stop, group);
      stops.push(stop);
    } else if (step.start === 'afterPrevious') {
      const previousEnd = groupEndMs(group!);
      if (previousEnd === null) return false;
      group = wrapperPar(ids.next++, String(previousEnd));
      appendInto(stop, group);
    }
    setNodeType(step.cTn, step.start);
    appendInto(group!, step.par);
  }

  layout.childTnLst.children = stops;
  if (stops.length === 0) {
    // CT_TimeNodeList requires a child, so an empty sequence drops the element
    // — and a `<p:seq>` with no time nodes left in it is only the navigation
    // scaffolding around nothing, so it goes too.
    const seq = mainSeqCTn(timing)!;
    seq.children = seq.children.filter((c) => c !== layout.childTnLst);
    const rootList = rootChildTnLst(timing);
    if (rootList !== null) {
      rootList.children = rootList.children.filter(
        (c) => c.kind !== 'element' || firstChildElement(c, NAME_C_TN) !== seq,
      );
    }
  }
  return true;
};

/**
 * Renumbers every `<p:cTn id>` under `el` by `offset`, so a freshly built
 * effect can be dropped into a tree without colliding with the ids in it.
 */
export const shiftCTnIds = (el: XmlElement, offset: number): void => {
  const walk = (e: XmlElement): void => {
    if (isPml(e, 'cTn')) {
      const raw = getAttrValue(e, ATTR_ID);
      if (raw !== null && /^\d+$/.test(raw)) {
        e.attrs = e.attrs.map((a) =>
          a.name.namespaceURI === '' && a.name.localName === 'id'
            ? { ...a, value: String(Number(raw) + offset) }
            : a,
        );
      }
    }
    for (const c of e.children) if (c.kind === 'element') walk(c);
  };
  walk(el);
};

/** The `spid|grpId` a `<p:bldP>` or an effect belongs to. */
const buildKey = (spid: string | null, grpId: string | null): string =>
  `${spid ?? ''}|${grpId ?? ''}`;

const effectBuildKeys = (timing: XmlElement): Set<string> => {
  const keys = new Set<string>();
  const walk = (el: XmlElement, grpId: string | null): void => {
    const own = isPml(el, 'cTn') ? getAttrValue(el, ATTR_GRP_ID) : null;
    const current = own ?? grpId;
    if (isPml(el, 'spTgt')) keys.add(buildKey(getAttrValue(el, ATTR_SPID), current));
    for (const c of el.children) if (c.kind === 'element') walk(c, current);
  };
  walk(timing, null);
  return keys;
};

/**
 * Drops the `<p:bldP>` entries in `keys` that nothing animates any more, so a
 * removed effect leaves no build pointing at a group that is gone. Entries that
 * were already orphaned before the edit are left alone: they are the file's,
 * not ours to tidy.
 */
export const pruneBuildEntries = (timing: XmlElement, keys: ReadonlySet<string>): void => {
  const bldLst = firstChildElement(timing, NAME_BLD_LST);
  if (bldLst === null || keys.size === 0) return;
  const live = effectBuildKeys(timing);
  bldLst.children = bldLst.children.filter((c) => {
    if (c.kind !== 'element' || !isPml(c, 'bldP')) return true;
    const key = buildKey(getAttrValue(c, ATTR_SPID), getAttrValue(c, ATTR_GRP_ID));
    return !keys.has(key) || live.has(key);
  });
  if (elementChildren(bldLst).length === 0) {
    timing.children = timing.children.filter((c) => c !== bldLst);
  }
};

/** The build entry an effect node belongs to, for pruning once it is gone. */
export const buildKeyOf = (effectCTn: XmlElement, spid: number): string =>
  buildKey(String(spid), getAttrValue(effectCTn, ATTR_GRP_ID));

/**
 * Turns a build entry's paragraph-by-paragraph reveal on or off.
 * `ST_TLParaBuildType` defaults to `whole`, so turning it off is the absence of
 * the attribute rather than a token of its own.
 */
export const setBuildByParagraph = (timing: XmlElement, key: string, on: boolean): void => {
  const bldLst = firstChildElement(timing, NAME_BLD_LST);
  if (bldLst === null) return;
  for (const child of elementChildren(bldLst)) {
    if (!isPml(child, 'bldP')) continue;
    if (buildKey(getAttrValue(child, ATTR_SPID), getAttrValue(child, ATTR_GRP_ID)) !== key)
      continue;
    const without = child.attrs.filter(
      (a) => !(a.name.namespaceURI === '' && a.name.localName === 'build'),
    );
    child.attrs = on ? [...without, attr(qname('', 'build', ''), 'p')] : without;
  }
};

/**
 * Attributes and behaviours this library writes on an effect. Replacing an
 * effect node — which is what changing its preset or its paragraph build does —
 * throws away whatever else is on it, so a node carrying anything outside this
 * is refused instead.
 */
const EFFECT_ATTRS = new Set([
  'id',
  'presetID',
  'presetClass',
  'presetSubtype',
  'fill',
  'grpId',
  'nodeType',
]);
const EFFECT_BEHAVIOURS = new Set(['set', 'anim']);

export const isPlainEffect = (par: XmlElement): boolean => {
  if (par.attrs.length > 0 || elementChildren(par).length !== 1) return false;
  const cTn = firstChildElement(par, NAME_C_TN);
  if (cTn === null) return false;
  if (!cTn.attrs.every((a) => a.name.namespaceURI === '' && EFFECT_ATTRS.has(a.name.localName))) {
    return false;
  }
  if (!elementChildren(cTn).every((c) => isPml(c, 'stCondLst') || isPml(c, 'childTnLst'))) {
    return false;
  }
  const childTnLst = firstChildElement(cTn, NAME_CHILD_TN_LST);
  return elementChildren(childTnLst).every(
    (c) => c.name.namespaceURI === NS.pml && EFFECT_BEHAVIOURS.has(c.name.localName),
  );
};

/** Removes the `<p:par>` wrappers whose own `<p:cTn>` is one of `doomed`. */
const dropEffectPars = (el: XmlElement, doomed: ReadonlySet<XmlElement>): void => {
  el.children = el.children.filter((c) => {
    if (c.kind !== 'element' || !isPml(c, 'par')) return true;
    const cTn = firstChildElement(c, NAME_C_TN);
    return cTn === null || !doomed.has(cTn);
  });
  for (const c of el.children) if (c.kind === 'element') dropEffectPars(c, doomed);
};

/**
 * Removes the time nodes left holding nothing, innermost first: a group with no
 * effects, a stop with no groups, a sequence with no stops. `CT_TimeNodeList`
 * requires a child, so an emptied `<p:childTnLst>` goes with them rather than
 * being written out empty.
 */
const pruneEmpty = (el: XmlElement): void => {
  for (const c of el.children) if (c.kind === 'element') pruneEmpty(c);
  el.children = el.children.filter((c) => {
    if (c.kind !== 'element') return true;
    if (isPml(c, 'childTnLst')) return elementChildren(c).length > 0;
    // A `<p:par>` / `<p:seq>` whose list of children is gone holds nothing.
    if (!isPml(c, 'par') && !isPml(c, 'seq')) return true;
    const cTn = firstChildElement(c, NAME_C_TN);
    return cTn !== null && firstChildElement(cTn, NAME_CHILD_TN_LST) !== null;
  });
};

/**
 * Takes every animation of `spids` out of a timing tree, along with the build
 * entries that named them, and retimes what is left.
 *
 * Effects outside the main sequence go too: an interactive sequence bound to a
 * shape that no longer exists would leave the file pointing at nothing. Media
 * time nodes are somebody else's business and are left alone.
 *
 * Nothing cascades: an effect on another shape is never removed because this
 * one went. The caller checks first that no such effect depends on these, and
 * refuses the whole deletion if one does.
 *
 * Returns false when nothing referred to those shapes. Throws when what is
 * left cannot be retimed, rather than leaving an effect waiting on a moment
 * that no longer happens.
 */
export const dropShapesFromTiming = (
  timing: XmlElement,
  spids: ReadonlySet<number>,
  doomed: ReadonlySet<XmlElement>,
  fn: string,
): boolean => {
  const bldLst = firstChildElement(timing, NAME_BLD_LST);
  const staleBuilds =
    bldLst !== null &&
    elementChildren(bldLst).some(
      (c) => isPml(c, 'bldP') && spids.has(Number(getAttrValue(c, ATTR_SPID))),
    );
  if (doomed.size === 0 && !staleBuilds) return false;

  // Retimed from the layout as it reads *now*: an effect that started when a
  // removed one finished has to move up, and once the removal has happened the
  // offsets left behind no longer describe a sequence this library could have
  // written, so reading them afterwards would find nothing to recompute.
  // A sequence that does not read as ours in the first place keeps the offsets
  // it has — the effects that are gone are gone either way.
  const layout = readMainSeqLayout(timing);
  if (layout !== null) {
    const surviving = layout.steps.filter((step) => !doomed.has(step.cTn));
    if (!writeMainSeqLayout(timing, layout, surviving, idsAfter(timing))) {
      throw new Error(
        `${fn}: an animation that is staying starts when one being removed ends, and this library ` +
          'cannot work out when the effect before it now finishes. Remove that animation first.',
      );
    }
  }

  dropEffectPars(timing, doomed);
  pruneEmpty(timing);
  if (bldLst !== null) {
    bldLst.children = bldLst.children.filter(
      (c) =>
        !(
          c.kind === 'element' &&
          isPml(c, 'bldP') &&
          spids.has(Number(getAttrValue(c, ATTR_SPID)))
        ),
    );
    if (elementChildren(bldLst).length === 0) {
      timing.children = timing.children.filter((c) => c !== bldLst);
    }
  }
  return true;
};

/** True once the timing tree holds no time nodes at all. */
export const isEmptyTiming = (timing: XmlElement): boolean => {
  const rootList = rootChildTnLst(timing);
  return rootList === null || elementChildren(rootList).length === 0;
};

/** The start condition an effect's `nodeType` names, or `null` for one we do not model. */
export const startOfEffect = (cTn: XmlElement): AnimationStartCondition | null =>
  START_OF_NODE_TYPE[getAttrValue(cTn, ATTR_NODE_TYPE) ?? ''] ?? null;

/** Every `<p:spTgt spid>` value at or under `el`. */
export const targetedSpids = (el: XmlElement): Set<string> => {
  const out = new Set<string>();
  const walk = (e: XmlElement): void => {
    if (isPml(e, 'spTgt')) {
      const spid = getAttrValue(e, ATTR_SPID);
      if (spid !== null) out.add(spid);
    }
    for (const c of e.children) if (c.kind === 'element') walk(c);
  };
  walk(el);
  return out;
};

/** Rewrites every `<p:spTgt spid>` at or under `el` through `idMap`. */
export const remapSpids = (el: XmlElement, idMap: ReadonlyMap<string, string>): void => {
  const walk = (e: XmlElement): void => {
    if (isPml(e, 'spTgt')) {
      e.attrs = e.attrs.map((a) =>
        a.name.namespaceURI === '' && a.name.localName === 'spid'
          ? { ...a, value: idMap.get(a.value) ?? a.value }
          : a,
      );
    }
    for (const c of e.children) if (c.kind === 'element') walk(c);
  };
  walk(el);
};

/**
 * Every `<p:bldP>` in the tree, keyed by the shape and build group it ties
 * together, so a caller walking many effects looks each one up instead of
 * searching the list again per effect.
 */
export const readBuildEntries = (timing: XmlElement): Map<string, XmlElement> => {
  const out = new Map<string, XmlElement>();
  const bldLst = firstChildElement(timing, NAME_BLD_LST);
  if (bldLst === null) return out;
  for (const child of elementChildren(bldLst)) {
    if (!isPml(child, 'bldP')) continue;
    out.set(buildKey(getAttrValue(child, ATTR_SPID), getAttrValue(child, ATTR_GRP_ID)), child);
  }
  return out;
};

/** The key `readBuildEntries` files a shape's build under. */
export const buildEntryKey = (spid: string, grpId: string | null): string => buildKey(spid, grpId);
