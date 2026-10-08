// Slide transitions.

import {
  NS,
  type XmlElement,
  type XmlNode,
  attr,
  cloneElement,
  elem,
  firstChildElement,
  getAttrValue,
  qname,
  walkElements,
} from '../../internal/xml/index.ts';
import {
  REL_TYPES,
  type SlideTransition,
  type TransitionOptions,
  buildTransition,
  transitionEffectNamespace,
} from '../../internal/presentationml/index.ts';
import {
  contentTypeForAudioFormat,
  detectAudioFormat,
  emptyRels,
  nextRelId,
} from '../../internal/opc/index.ts';
import {
  INTERNAL_PACKAGE,
  SLIDE_DOCUMENT,
  SLIDE_PART_NAME,
  type SlideData,
} from '../_internal-symbols.ts';
import { commitSlideData, refreshSlideData } from './_helpers.ts';
import { internMediaPart } from './media.ts';

const isPml = (node: XmlNode, localName: string): boolean =>
  node.kind === 'element' && node.name.namespaceURI === NS.pml && node.name.localName === localName;

// The reference desktop app (2010 and later) writes a transition with a duration (or a
// newer effect) as `mc:AlternateContent`: the p14 choice first, an ECMA-376 fallback second.
// The choice is the one the reference desktop app itself reads.
const transitionsIn = (node: XmlNode): XmlElement[] => {
  if (node.kind !== 'element') return [];
  if (isPml(node, 'transition')) return [node];
  if (node.name.namespaceURI !== NS.mc || node.name.localName !== 'AlternateContent') return [];
  return node.children.flatMap((branch) =>
    branch.kind === 'element' && branch.name.namespaceURI === NS.mc
      ? branch.children.filter(
          (child): child is XmlElement => child.kind === 'element' && isPml(child, 'transition'),
        )
      : [],
  );
};
const transitionIn = (node: XmlNode): XmlElement | null => transitionsIn(node)[0] ?? null;

const findTransition = (slide: SlideData): XmlElement | null => {
  for (const node of slide[SLIDE_DOCUMENT].root.children) {
    const found = transitionIn(node);
    if (found) return found;
  }
  return null;
};

const removeTransition = (slide: SlideData): void => {
  slide[SLIDE_DOCUMENT].root.children = slide[SLIDE_DOCUMENT].root.children.filter(
    (c) => transitionIn(c) === null,
  );
};

const insertAfterClrMapOvr = (slide: SlideData, t: XmlElement): void => {
  const children = slide[SLIDE_DOCUMENT].root.children;
  let insertAt = children.length;
  for (let i = 0; i < children.length; i++) {
    const c = children[i];
    if (c?.kind !== 'element' || c.name.namespaceURI !== NS.pml) continue;
    if (c.name.localName === 'clrMapOvr') {
      insertAt = i + 1;
    } else if (c.name.localName === 'cSld' && insertAt === children.length) {
      insertAt = i + 1;
    }
  }
  children.splice(insertAt, 0, t);
};

const xmlBoolean = (el: XmlElement, name: string): boolean | undefined => {
  const raw = getAttrValue(el, qname('', name, ''))?.trim();
  return raw === undefined ? undefined : raw === '1' || raw === 'true';
};

type EffectFields = Omit<
  SlideTransition,
  'speed' | 'advanceOnClick' | 'advanceAfterMs' | 'durationMs'
>;

/**
 * The effect element and its options. The first child that is neither the
 * sound nor the extension list is the effect (CT_SlideTransition's choice).
 * An element this library does not write — another vendor's, or a token in a
 * namespace it does not belong to — is reported by its prefixed name, so it can
 * never pass for one of ours and be rewritten as such.
 */
const readEffect = (transition: XmlElement): EffectFields => {
  const child = transition.children.find(
    (c): c is XmlElement => c.kind === 'element' && !isPml(c, 'sndAc') && !isPml(c, 'extLst'),
  );
  if (child === undefined) return { effect: 'none' };
  const local = child.name.localName;
  if (transitionEffectNamespace(local) !== child.name.namespaceURI) {
    return { effect: child.name.prefix === '' ? local : `${child.name.prefix}:${local}` };
  }
  const text = (name: string): string | null => getAttrValue(child, qname('', name, ''));
  const direction = text('dir');
  const orientation = text('orient');
  const spokes = text('spokes');
  const pattern = text('pattern');
  const preset = text('prst');
  const option = text('option');
  const thruBlack = xmlBoolean(child, 'thruBlk');
  const isContent = xmlBoolean(child, 'isContent');
  const isInverted = xmlBoolean(child, 'isInverted');
  const hasBounce = xmlBoolean(child, 'hasBounce');
  const invertX = xmlBoolean(child, 'invX');
  const invertY = xmlBoolean(child, 'invY');
  return {
    effect: local,
    ...(direction !== null ? { direction } : {}),
    ...(orientation === 'horz' || orientation === 'vert' ? { orientation } : {}),
    ...((local === 'wheel' || local === 'wheelReverse') && spokes !== null
      ? { spokes: Number(spokes) }
      : {}),
    ...(thruBlack !== undefined ? { thruBlack } : {}),
    ...(pattern === 'diamond' ||
    pattern === 'hexagon' ||
    pattern === 'strip' ||
    pattern === 'rectangle'
      ? { pattern }
      : {}),
    ...(isContent !== undefined ? { isContent } : {}),
    ...(isInverted !== undefined ? { isInverted } : {}),
    ...(hasBounce !== undefined ? { hasBounce } : {}),
    ...(preset !== null ? { preset } : {}),
    ...(invertX !== undefined ? { invertX } : {}),
    ...(invertY !== undefined ? { invertY } : {}),
    ...(option === 'byObject' || option === 'byWord' || option === 'byChar'
      ? { morphOption: option }
      : {}),
  };
};

/**
 * Reads back the slide's current transition (or `null` if no
 * `<p:transition>` is present). The returned shape mirrors what
 * `setSlideTransition` accepts.
 */
export const getSlideTransition = (slide: SlideData): SlideTransition | null => {
  const transition = findTransition(slide);
  if (!transition) return null;
  const duration = getAttrValue(transition, qname('p14', 'dur', NS.p14));
  const speed = getAttrValue(transition, qname('', 'spd', '')) as 'slow' | 'med' | 'fast' | null;
  const advClick = getAttrValue(transition, qname('', 'advClick', ''))?.trim() ?? null;
  const advTm = getAttrValue(transition, qname('', 'advTm', ''));
  return {
    ...readEffect(transition),
    ...(speed !== null ? { speed } : {}),
    ...(advClick !== null ? { advanceOnClick: advClick !== '0' && advClick !== 'false' } : {}),
    ...(advTm !== null ? { advanceAfterMs: Number.parseInt(advTm, 10) } : {}),
    ...(duration !== null ? { durationMs: Number.parseInt(duration, 10) } : {}),
  };
};

/**
 * Sets the slide's transition effect. A sound set with
 * `setSlideTransitionSound` is kept, as the reference desktop app keeps it when the effect
 * changes.
 */
export const setSlideTransition = (slide: SlideData, options: TransitionOptions): void => {
  const existing = findTransition(slide);
  const sound = existing ? firstChildElement(existing, NAME_SND_AC) : null;
  const transition = buildTransition(options, sound);
  removeTransition(slide);
  insertAfterClrMapOvr(slide, transition);
  commitSlideData(slide);
  refreshSlideData(slide);
};

/** Removes any existing transition on the slide, including its sound. */
export const clearSlideTransition = (slide: SlideData): void => {
  removeTransition(slide);
  releaseUnusedSoundRels(slide);
  commitSlideData(slide);
  refreshSlideData(slide);
};

const NAME_SND_AC = qname('p', 'sndAc', NS.pml);
const NAME_ST_SND = qname('p', 'stSnd', NS.pml);
const NAME_END_SND = qname('p', 'endSnd', NS.pml);
const NAME_SND = qname('p', 'snd', NS.pml);
const ATTR_EMBED = qname('r', 'embed', NS.officeDocRels);
const ATTR_NAME = qname('', 'name', '');
const ATTR_LOOP = qname('', 'loop', '');

/** What plays when the slide's transition starts. */
export type TransitionSound =
  | { readonly kind: 'stop' }
  | { readonly kind: 'play'; readonly name: string; readonly loop: boolean };

/**
 * `stop` ends any sound still playing from an earlier slide ("Stop Previous
 * Sound"); `play` embeds a WAV file, the only format `p:snd` allows
 * (ECMA-376 §19.5.66, CT_EmbeddedWAVAudioFile).
 */
export type TransitionSoundInput =
  | { readonly kind: 'stop' }
  | {
      readonly kind: 'play';
      readonly data: Uint8Array;
      readonly name: string;
      readonly loop?: boolean;
    };

/** Reads the transition's sound, or `null` when it has none. */
export const getSlideTransitionSound = (slide: SlideData): TransitionSound | null => {
  const transition = findTransition(slide);
  const action = transition ? firstChildElement(transition, NAME_SND_AC) : null;
  if (!action) return null;
  if (firstChildElement(action, NAME_END_SND)) return { kind: 'stop' };
  const start = firstChildElement(action, NAME_ST_SND);
  const snd = start ? firstChildElement(start, NAME_SND) : null;
  if (!start || !snd) return null;
  const loop = getAttrValue(start, ATTR_LOOP);
  return {
    kind: 'play',
    name: getAttrValue(snd, ATTR_NAME) ?? '',
    loop: loop === '1' || loop === 'true',
  };
};

// The `audio` rel of a transition sound is owned by its `p:snd`; once nothing
// in the slide references it, it only keeps the clip in the file.
const releaseUnusedSoundRels = (slide: SlideData): void => {
  const pkg = slide[INTERNAL_PACKAGE];
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (rels === null) return;
  const referenced = new Set<string>();
  walkElements(slide[SLIDE_DOCUMENT].root, (element) => {
    for (const a of element.attrs)
      if (a.name.namespaceURI === NS.officeDocRels) referenced.add(a.value);
  });
  const kept = rels.items.filter((rel) => rel.type !== REL_TYPES.audio || referenced.has(rel.id));
  if (kept.length !== rels.items.length) pkg.setRels(slide[SLIDE_PART_NAME], { items: kept });
};

/**
 * Sets (or with `null` removes) the sound played when the slide's transition
 * starts. A slide without a transition gets one with no effect, which is how
 * the reference desktop app stores a sound on a "None" transition.
 */
export const setSlideTransitionSound = (
  slide: SlideData,
  sound: TransitionSoundInput | null,
): void => {
  if (sound?.kind === 'play' && detectAudioFormat(sound.data) !== 'wav')
    throw new Error('setSlideTransitionSound: a transition sound must be a WAV file');
  let transition = findTransition(slide);
  if (!transition) {
    if (sound === null) return;
    transition = buildTransition({ effect: 'none' });
    insertAfterClrMapOvr(slide, transition);
  }
  // Every copy (the p14 choice and its fallback) carries the same sound.
  const copies = slide[SLIDE_DOCUMENT].root.children.flatMap((node) => {
    const found = transitionIn(node);
    if (!found) return [];
    if (found === node) return [found];
    return (node as XmlElement).children.flatMap((branch) =>
      branch.kind === 'element'
        ? branch.children.filter((child) => isPml(child, 'transition'))
        : [],
    ) as XmlElement[];
  });

  let action: XmlElement | null = null;
  if (sound?.kind === 'stop') action = elem(NAME_SND_AC, { children: [elem(NAME_END_SND)] });
  else if (sound?.kind === 'play') {
    const pkg = slide[INTERNAL_PACKAGE];
    const part = internMediaPart(pkg, 'wav', contentTypeForAudioFormat('wav'), sound.data);
    const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
    const target = `../media/${part.slice(part.lastIndexOf('/') + 1)}`;
    let id = rels.items.find((rel) => rel.type === REL_TYPES.audio && rel.target === target)?.id;
    if (id === undefined) {
      id = nextRelId(rels.items.map((rel) => rel.id));
      rels.items.push({ id, type: REL_TYPES.audio, target, targetMode: 'Internal' });
      pkg.setRels(slide[SLIDE_PART_NAME], rels);
    }
    action = elem(NAME_SND_AC, {
      children: [
        elem(NAME_ST_SND, {
          attrs: sound.loop ? [attr(ATTR_LOOP, '1')] : [],
          children: [
            elem(NAME_SND, { attrs: [attr(ATTR_EMBED, id), attr(ATTR_NAME, sound.name)] }),
          ],
        }),
      ],
    });
  }
  for (const copy of copies) {
    // CT_SlideTransition: the effect choice, then sndAc, then extLst.
    copy.children = copy.children.filter((child) => !isPml(child, 'sndAc'));
    if (action === null) continue;
    const extLst = copy.children.findIndex((child) => isPml(child, 'extLst'));
    copy.children.splice(extLst === -1 ? copy.children.length : extLst, 0, cloneElement(action));
  }
  releaseUnusedSoundRels(slide);
  commitSlideData(slide);
  refreshSlideData(slide);
};
