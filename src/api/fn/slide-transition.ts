// Slide transitions.

import { NS, type XmlElement, getAttrValue, qname } from '../../internal/xml/index.ts';
import { type TransitionOptions, buildTransition } from '../../internal/presentationml/index.ts';
import { SLIDE_DOCUMENT, type SlideData } from '../_internal-symbols.ts';
import { commitSlideData, refreshSlideData } from './_helpers.ts';

const removeTransition = (slide: SlideData): void => {
  slide[SLIDE_DOCUMENT].root.children = slide[SLIDE_DOCUMENT].root.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.pml &&
        c.name.localName === 'transition'
      ),
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

/**
 * Reads back the slide's current transition (or `null` if no
 * `<p:transition>` is present). The returned shape mirrors what
 * `setSlideTransition` accepts.
 */
export const getSlideTransition = (slide: SlideData): TransitionOptions | null => {
  const transition = slide[SLIDE_DOCUMENT].root.children.find(
    (c): c is XmlElement =>
      c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'transition',
  );
  if (!transition) return null;
  const speed = getAttrValue(transition, qname('', 'spd', '')) as 'slow' | 'med' | 'fast' | null;
  const advClick = getAttrValue(transition, qname('', 'advClick', ''));
  const advTm = getAttrValue(transition, qname('', 'advTm', ''));
  // First child element identifies the effect (`p:fade`, `p:wipe`, ...).
  let effect: string | null = null;
  let direction: string | null = null;
  let orientation: 'horz' | 'vert' | null = null;
  let thruBlack: boolean | undefined;
  for (const child of transition.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.pml) continue;
    effect = child.name.localName;
    direction = getAttrValue(child, qname('', 'dir', ''));
    const o = getAttrValue(child, qname('', 'orient', ''));
    if (o === 'horz' || o === 'vert') orientation = o;
    const tb = getAttrValue(child, qname('', 'thruBlk', ''));
    if (tb !== null) thruBlack = tb === '1';
    break;
  }
  if (effect === null) return null;
  return {
    effect,
    ...(speed !== null ? { speed } : {}),
    ...(direction !== null ? { direction } : {}),
    ...(orientation !== null ? { orientation } : {}),
    ...(thruBlack !== undefined ? { thruBlack } : {}),
    ...(advClick !== null ? { advanceOnClick: advClick !== '0' } : {}),
    ...(advTm !== null ? { advanceAfterMs: Number.parseInt(advTm, 10) } : {}),
  };
};

/** Sets the slide's transition effect. */
export const setSlideTransition = (slide: SlideData, options: TransitionOptions): void => {
  removeTransition(slide);
  insertAfterClrMapOvr(slide, buildTransition(options));
  commitSlideData(slide);
  refreshSlideData(slide);
};

/** Removes any existing transition on the slide. */
export const clearSlideTransition = (slide: SlideData): void => {
  removeTransition(slide);
  commitSlideData(slide);
  refreshSlideData(slide);
};
