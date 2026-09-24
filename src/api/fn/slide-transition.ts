// Slide transitions.

import {
  NS,
  attr,
  elem,
  firstChildElement,
  type XmlElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';
import { type TransitionOptions, buildTransition } from '../../internal/presentationml/index.ts';
import {
  INTERNAL_PACKAGE,
  SLIDE_PART_NAME,
  SLIDE_DOCUMENT,
  type PresentationData,
  type SlideData,
} from '../_internal-symbols.ts';
import {
  detectAudioFormat,
  emptyRels,
  nextRelId,
  partName,
  resolveTarget,
} from '../../internal/opc/index.ts';
import { getSlides } from './slide-query.ts';
import { commitSlideData, refreshSlideData } from './_helpers.ts';

const isTransition = (node: XmlElement): boolean =>
  node.name.namespaceURI === NS.pml && node.name.localName === 'transition';

// PowerPoint stores modern transitions in Choice/Fallback branches.
const transitionsIn = (node: XmlElement): XmlElement[] => {
  if (isTransition(node)) return [node];
  if (
    node.name.namespaceURI !== NS.mc ||
    !['AlternateContent', 'Choice', 'Fallback'].includes(node.name.localName)
  )
    return [];
  return node.children.flatMap((child) => (child.kind === 'element' ? transitionsIn(child) : []));
};
const transitionContainer = (slide: SlideData): XmlElement | undefined =>
  slide[SLIDE_DOCUMENT].root.children.find(
    (child): child is XmlElement => child.kind === 'element' && transitionsIn(child).length > 0,
  );
const transitionNodes = (slide: SlideData): XmlElement[] => {
  const container = transitionContainer(slide);
  return container ? transitionsIn(container) : [];
};
// Read only a usable compatibility branch. Keep transitionNodes() for mutations:
// edits must continue updating every stored alternative without discarding it.
const readableTransition = (slide: SlideData): XmlElement | undefined => {
  const container = transitionContainer(slide);
  if (!container) return undefined;
  const read = (node: XmlElement, inherited: Map<string, string>): XmlElement | undefined => {
    const namespaces = new Map([...inherited, ...node.prefixDecls]);
    if (isTransition(node)) return node;
    if (node.name.namespaceURI !== NS.mc) return undefined;
    const children = node.children.filter((child): child is XmlElement => child.kind === 'element');
    if (node.name.localName === 'AlternateContent') {
      for (const choice of children) {
        if (choice.name.namespaceURI !== NS.mc || choice.name.localName !== 'Choice') continue;
        const scope = new Map([...namespaces, ...choice.prefixDecls]);
        const requires = (getAttrValue(choice, qname('', 'Requires', '')) ?? '')
          .trim()
          .split(/\s+/);
        if (
          !requires.every((prefix) =>
            new Set<string>([NS.pml, NS.p14]).has(scope.get(prefix) ?? ''),
          )
        )
          continue;
        const transition = read(choice, namespaces);
        // p14:dur is supported; p14 visual effects still need their fallback.
        if (
          transition &&
          transition.children.every(
            (child) =>
              child.kind !== 'element' ||
              child.name.namespaceURI === NS.pml ||
              (child.name.namespaceURI === NS.p14 && child.name.localName === 'wheelReverse'),
          )
        )
          return transition;
      }
      const fallback = children.find(
        (child) => child.name.namespaceURI === NS.mc && child.name.localName === 'Fallback',
      );
      return fallback ? read(fallback, namespaces) : undefined;
    }
    if (!['Choice', 'Fallback'].includes(node.name.localName)) return undefined;
    for (const child of children) {
      const transition = read(child, namespaces);
      if (transition) return transition;
    }
    return undefined;
  };
  return read(container, slide[SLIDE_DOCUMENT].root.prefixDecls);
};

const removeTransition = (slide: SlideData): void => {
  slide[SLIDE_DOCUMENT].root.children = slide[SLIDE_DOCUMENT].root.children.filter(
    (child) => child.kind !== 'element' || !transitionsIn(child).length,
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

export type SlideTransitionSound =
  | { readonly kind: 'none' }
  | { readonly kind: 'stop' }
  | {
      readonly kind: 'play';
      readonly name: string;
      readonly bytes: Uint8Array;
      readonly loop: boolean;
    };

/** Reads an embedded transition sound. Missing or external audio is not loaded. */
export const getSlideTransitionSound = (slide: SlideData): SlideTransitionSound => {
  const transition = readableTransition(slide);
  const action = transition && firstChildElement(transition, qname('p', 'sndAc', NS.pml));
  if (!action) return { kind: 'none' };
  if (firstChildElement(action, qname('p', 'endSnd', NS.pml))) return { kind: 'stop' };
  const start = firstChildElement(action, qname('p', 'stSnd', NS.pml));
  const sound = start && firstChildElement(start, qname('p', 'snd', NS.pml));
  if (!sound || !start) return { kind: 'none' };
  const id = getAttrValue(sound, qname('r', 'embed', NS.officeDocRels));
  const pkg = slide[INTERNAL_PACKAGE];
  const relationship = pkg
    .getRels(slide[SLIDE_PART_NAME])
    ?.items.find(
      (rel) =>
        rel.id === id && rel.type === NS.officeDocRels + '/audio' && rel.targetMode !== 'External',
    );
  const media =
    relationship && pkg.getPart(resolveTarget(slide[SLIDE_PART_NAME], relationship.target));
  if (!media) return { kind: 'none' };
  return {
    kind: 'play',
    name: getAttrValue(sound, qname('', 'name', '')) ?? '',
    bytes: media.data.slice(),
    loop: ['1', 'true'].includes(getAttrValue(start, qname('', 'loop', '')) ?? ''),
  };
};

/** Sets only transition audio, retaining effects/timing and shared media parts. */
export const setSlideTransitionSound = (slide: SlideData, value: SlideTransitionSound): void => {
  if (!value || !['none', 'stop', 'play'].includes(value.kind))
    throw new Error('Invalid transition sound.');
  if (
    value.kind === 'play' &&
    (typeof value.name !== 'string' ||
      typeof value.loop !== 'boolean' ||
      !(value.bytes instanceof Uint8Array) ||
      detectAudioFormat(value.bytes) !== 'wav')
  )
    throw new Error('Transition sound must be a WAV file with a name and loop setting.');
  const pkg = slide[INTERNAL_PACKAGE];
  let action: XmlElement | undefined;
  if (value.kind === 'stop')
    action = elem(qname('p', 'sndAc', NS.pml), { children: [elem(qname('p', 'endSnd', NS.pml))] });
  if (value.kind === 'play') {
    let media = pkg.parts.find(
      (part) =>
        part.contentType.toLowerCase().includes('wav') &&
        part.data.length === value.bytes.length &&
        part.data.every((byte, index) => byte === value.bytes[index]),
    );
    if (!media) {
      let ordinal = 1;
      while (pkg.getPart(partName('/ppt/media/transitionSound' + ordinal + '.wav'))) ordinal++;
      media = pkg.addPart(
        partName('/ppt/media/transitionSound' + ordinal + '.wav'),
        'audio/x-wav',
        value.bytes.slice(),
      );
    }
    const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
    const type = NS.officeDocRels + '/audio';
    const existing = rels.items.find(
      (rel) =>
        rel.type === type &&
        rel.targetMode !== 'External' &&
        resolveTarget(slide[SLIDE_PART_NAME], rel.target) === media.name,
    );
    const id = existing?.id ?? nextRelId(rels.items.map((rel) => rel.id));
    if (!existing) {
      rels.items.push({ id, type, target: media.name, targetMode: 'Internal' });
      pkg.setRels(slide[SLIDE_PART_NAME], rels);
    }
    const sound = elem(qname('p', 'snd', NS.pml), {
      attrs: [
        attr(qname('r', 'embed', NS.officeDocRels), id),
        attr(qname('', 'name', ''), value.name),
      ],
    });
    sound.prefixDecls.set('r', NS.officeDocRels);
    action = elem(qname('p', 'sndAc', NS.pml), {
      children: [
        elem(qname('p', 'stSnd', NS.pml), {
          attrs: [attr(qname('', 'loop', ''), value.loop ? '1' : '0')],
          children: [sound],
        }),
      ],
    });
  }
  const nodes = transitionNodes(slide);
  if (!nodes.length && action) {
    const transition = buildTransition({ effect: 'none' });
    insertAfterClrMapOvr(slide, transition);
    nodes.push(transition);
  }
  for (const transition of nodes) {
    transition.children = transition.children.filter(
      (child) =>
        !(
          child.kind === 'element' &&
          child.name.namespaceURI === NS.pml &&
          child.name.localName === 'sndAc'
        ),
    );
    if (action) {
      const extension = transition.children.findIndex(
        (child) =>
          child.kind === 'element' &&
          child.name.namespaceURI === NS.pml &&
          child.name.localName === 'extLst',
      );
      transition.children.splice(
        extension < 0 ? transition.children.length : extension,
        0,
        structuredClone(action),
      );
    }
  }
  commitSlideData(slide);
  refreshSlideData(slide);
};

/**
 * Reads the first supported compatibility alternative (or `null` if none
 * is available). Modern effects use their stored fallback. The returned shape mirrors what
 * `setSlideTransition` accepts.
 */
export const getSlideTransition = (slide: SlideData): TransitionOptions | null => {
  const transition = readableTransition(slide);
  if (!transition) return null;
  const speed = getAttrValue(transition, qname('', 'spd', '')) as 'slow' | 'med' | 'fast' | null;
  const advClick = getAttrValue(transition, qname('', 'advClick', ''));
  const advTm = getAttrValue(transition, qname('', 'advTm', ''));
  const duration = getAttrValue(transition, qname('p14', 'dur', NS.p14));
  // Sound actions and extensions may occur without a visual effect.
  let effect: string | null = null;
  let direction: string | null = null;
  let orientation: 'horz' | 'vert' | null = null;
  let thruBlack: boolean | undefined;
  let spokes: number | undefined;
  for (const child of transition.children) {
    if (
      child.kind !== 'element' ||
      (child.name.namespaceURI !== NS.pml &&
        !(child.name.namespaceURI === NS.p14 && child.name.localName === 'wheelReverse'))
    )
      continue;
    if (child.name.localName === 'sndAc' || child.name.localName === 'extLst') continue;
    effect = child.name.localName;
    const spokeCount = getAttrValue(child, qname('', 'spokes', ''));
    if (['wheel', 'wheelReverse'].includes(effect) && spokeCount !== null)
      spokes = Number(spokeCount);
    direction = getAttrValue(child, qname('', 'dir', ''));
    const o = getAttrValue(child, qname('', 'orient', ''));
    if (o === 'horz' || o === 'vert') orientation = o;
    const tb = getAttrValue(child, qname('', 'thruBlk', ''));
    if (tb !== null) thruBlack = tb === '1' || tb === 'true';
    break;
  }
  return {
    effect: effect ?? 'none',
    ...(spokes !== undefined ? { spokes } : {}),
    ...(speed !== null ? { speed } : {}),
    ...(duration !== null ? { durationMs: Number.parseInt(duration, 10) } : {}),
    ...(direction !== null ? { direction } : {}),
    ...(orientation !== null ? { orientation } : {}),
    ...(thruBlack !== undefined ? { thruBlack } : {}),
    ...(advClick !== null ? { advanceOnClick: advClick === '1' || advClick === 'true' } : {}),
    ...(advTm !== null ? { advanceAfterMs: Number.parseInt(advTm, 10) } : {}),
  };
};

// Modern effects are stored with a legacy alternative for older readers.
const transitionWithFallback = (transition: XmlElement, effect: string): XmlElement => {
  if (effect !== 'wheelReverse') return transition;
  const fallback = structuredClone(transition);
  for (const child of fallback.children) {
    if (
      child.kind === 'element' &&
      child.name.namespaceURI === NS.p14 &&
      child.name.localName === 'wheelReverse'
    )
      child.name = qname('p', 'wheel', NS.pml);
  }
  return elem(qname('mc', 'AlternateContent', NS.mc), {
    prefixDecls: new Map([
      ['mc', NS.mc],
      ['p14', NS.p14],
    ]),
    children: [
      elem(qname('mc', 'Choice', NS.mc), {
        attrs: [attr(qname('', 'Requires', ''), 'p14')],
        children: [transition],
      }),
      elem(qname('mc', 'Fallback', NS.mc), { children: [fallback] }),
    ],
  });
};

/** Sets the slide's transition effect. */
export const setSlideTransition = (slide: SlideData, options: TransitionOptions): void => {
  const transition = buildTransition(options);
  removeTransition(slide);
  insertAfterClrMapOvr(slide, transitionWithFallback(transition, options.effect));
  commitSlideData(slide);
  refreshSlideData(slide);
};

/** Changes only the visual effect, retaining timing, sound and extension metadata. */
export const setSlideTransitionEffect = (
  slide: SlideData,
  options: Pick<TransitionOptions, 'effect' | 'direction' | 'orientation' | 'thruBlack' | 'spokes'>,
): void => {
  // Build first: invalid choices must not change the slide.
  const replacement = buildTransition(options);
  const nodes = transitionNodes(slide);
  if (options.effect === 'wheelReverse') {
    const current = readableTransition(slide) ?? nodes[0];
    const transition = current ? structuredClone(current) : replacement;
    // Carry namespace bindings used by preserved metadata into the new wrapper.
    const inherit = (node: XmlElement, bindings: Map<string, string>): boolean => {
      const scope = new Map([...bindings, ...node.prefixDecls]);
      if (node === current) {
        transition.prefixDecls = scope;
        return true;
      }
      return node.children.some((child) => child.kind === 'element' && inherit(child, scope));
    };
    if (current) inherit(slide[SLIDE_DOCUMENT].root, new Map());
    transition.children = [
      ...structuredClone(replacement.children),
      ...transition.children.filter(
        (child) =>
          child.kind !== 'element' ||
          (child.name.namespaceURI === NS.pml &&
            ['sndAc', 'extLst'].includes(child.name.localName)),
      ),
    ];
    removeTransition(slide);
    insertAfterClrMapOvr(slide, transitionWithFallback(transition, options.effect));
    commitSlideData(slide);
    refreshSlideData(slide);
    return;
  }
  if (!nodes.length) insertAfterClrMapOvr(slide, replacement);
  for (const transition of nodes) {
    transition.children = [
      ...structuredClone(replacement.children),
      ...transition.children.filter(
        (child) =>
          child.kind !== 'element' ||
          (child.name.namespaceURI === NS.pml &&
            ['sndAc', 'extLst'].includes(child.name.localName)),
      ),
    ];
  }
  commitSlideData(slide);
  refreshSlideData(slide);
};

/** Removes any existing transition on the slide. */
export const clearSlideTransition = (slide: SlideData): void => {
  removeTransition(slide);
  commitSlideData(slide);
  refreshSlideData(slide);
};

/** Changes slide advance timing without replacing effects, sounds or extension metadata. */
export const setSlideAdvanceTiming = (
  slide: SlideData,
  timing: { advanceOnClick: boolean; advanceAfterMs: number | null },
): void => {
  if (!timing || typeof timing.advanceOnClick !== 'boolean')
    throw new Error('Invalid advance-on-click setting.');
  const delay = timing.advanceAfterMs;
  if (delay !== null && (!Number.isInteger(delay) || delay < 0 || delay > 4294967295))
    throw new Error('Invalid slide advance time.');
  const nodes = transitionNodes(slide);
  if (!nodes.length) {
    const transition = buildTransition({ effect: 'none' });
    insertAfterClrMapOvr(slide, transition);
    nodes.push(transition);
  }
  for (const transition of nodes) {
    transition.attrs = transition.attrs.filter(
      (attribute) =>
        attribute.name.namespaceURI !== '' ||
        !['advClick', 'advTm'].includes(attribute.name.localName),
    );
    transition.attrs.push(attr(qname('', 'advClick', ''), timing.advanceOnClick ? '1' : '0'));
    if (delay !== null) transition.attrs.push(attr(qname('', 'advTm', ''), String(delay)));
  }
  commitSlideData(slide);
  refreshSlideData(slide);
};

/** Changes visual transition duration in milliseconds while preserving other settings. */
export const setSlideTransitionDuration = (slide: SlideData, durationMs: number): void => {
  const replacement = buildTransition({ effect: 'none', durationMs });
  const nodes = transitionNodes(slide);
  if (!nodes.length) insertAfterClrMapOvr(slide, replacement);
  for (const transition of nodes) {
    // Bind extension attributes locally, avoiding prefixes already in scope.
    const bindings = new Map([
      ...slide[SLIDE_DOCUMENT].root.prefixDecls,
      ...transition.prefixDecls,
    ]);
    let prefix = 'p14';
    while (bindings.has(prefix) && bindings.get(prefix) !== NS.p14) prefix += '_';
    transition.prefixDecls.set(prefix, NS.p14);
    transition.attrs = transition.attrs.filter(
      (a) => !(a.name.namespaceURI === NS.p14 && a.name.localName === 'dur'),
    );
    transition.attrs.push(attr(qname(prefix, 'dur', NS.p14), String(durationMs)));
    const ignorable = transition.attrs.find(
      (a) => a.name.namespaceURI === NS.mc && a.name.localName === 'Ignorable',
    );
    const tokens = new Set((ignorable?.value ?? '').split(/\s+/).filter(Boolean));
    tokens.add(prefix);
    if (ignorable)
      transition.attrs[transition.attrs.indexOf(ignorable)] = {
        ...ignorable,
        value: [...tokens].join(' '),
      };
    else {
      let mcPrefix = 'mc';
      while (bindings.has(mcPrefix) && bindings.get(mcPrefix) !== NS.mc) mcPrefix += '_';
      transition.prefixDecls.set(mcPrefix, NS.mc);
      transition.attrs.push(attr(qname(mcPrefix, 'Ignorable', NS.mc), [...tokens].join(' ')));
    }
  }
  commitSlideData(slide);
  refreshSlideData(slide);
};

/** Applies the source slide's complete transition to every slide in the same presentation. */
export const applySlideTransitionToAll = (
  presentation: PresentationData,
  source: SlideData,
): void => {
  const pkg = presentation[INTERNAL_PACKAGE];
  const slides = getSlides(presentation);
  if (
    source[INTERNAL_PACKAGE] !== pkg ||
    !slides.some((slide) => slide[SLIDE_PART_NAME] === source[SLIDE_PART_NAME])
  )
    throw new Error('The transition source does not belong to this presentation.');
  const original = transitionContainer(source);
  const sourceRels = pkg.getRels(source[SLIDE_PART_NAME]);
  // Prepare all copies first so invalid relationships cannot leave a partially changed deck.
  const changes = slides
    .filter((slide) => slide[SLIDE_PART_NAME] !== source[SLIDE_PART_NAME])
    .map((slide) => {
      const copy = original ? structuredClone(original) : null;
      const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
      let relsChanged = false;
      if (copy) {
        copy.prefixDecls = new Map([
          ...source[SLIDE_DOCUMENT].root.prefixDecls,
          ...copy.prefixDecls,
        ]);
        for (const inherited of source[SLIDE_DOCUMENT].root.attrs.filter(
          (a) => a.name.namespaceURI === NS.mc,
        )) {
          if (
            !copy.attrs.some(
              (a) => a.name.namespaceURI === NS.mc && a.name.localName === inherited.name.localName,
            )
          )
            copy.attrs.push(structuredClone(inherited));
        }
        const visit = (node: XmlElement) => {
          for (const attribute of node.attrs.filter(
            (a) => a.name.namespaceURI === NS.officeDocRels && a.value,
          )) {
            const relationship = sourceRels?.items.find((rel) => rel.id === attribute.value);
            if (!relationship)
              throw new Error('Transition relationship is missing: ' + attribute.value);
            const external = relationship.targetMode === 'External';
            const target = external
              ? relationship.target
              : resolveTarget(source[SLIDE_PART_NAME], relationship.target);
            if (!external && !pkg.getPart(partName(target)))
              throw new Error('Transition relationship target is missing.');
            const existing = rels.items.find(
              (rel) =>
                rel.type === relationship.type &&
                (rel.targetMode === 'External') === external &&
                (external ? rel.target : resolveTarget(slide[SLIDE_PART_NAME], rel.target)) ===
                  target,
            );
            if (existing)
              node.attrs[node.attrs.indexOf(attribute)] = { ...attribute, value: existing.id };
            else {
              const id = nextRelId(rels.items.map((rel) => rel.id));
              rels.items.push({ ...relationship, id, target });
              node.attrs[node.attrs.indexOf(attribute)] = { ...attribute, value: id };
              relsChanged = true;
            }
          }
          for (const child of node.children) if (child.kind === 'element') visit(child);
        };
        visit(copy);
      }
      return { slide, copy, rels, relsChanged };
    });
  for (const { slide, copy, rels, relsChanged } of changes) {
    if (relsChanged) pkg.setRels(slide[SLIDE_PART_NAME], rels);
    removeTransition(slide);
    if (copy) insertAfterClrMapOvr(slide, copy);
    commitSlideData(slide);
    refreshSlideData(slide);
  }
};
