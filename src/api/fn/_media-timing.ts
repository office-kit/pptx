// Media time nodes (`<p:video>` / `<p:audio>`) inside a slide's `<p:timing>`.
//
// A clip's play button / seek bar only appears in the slide show when the
// timing tree's root holds a media node targeting the picture, so the shape
// and its node have to be added, copied and removed together. The animation
// code shares the same `<p:timing>` element; these helpers touch only the
// media nodes and leave every animation sequence as they found it.

import {
  type MediaFileKind,
  buildMediaTimingNode,
  buildTimingRoot,
  isMediaTimingNode,
  mediaTimingNodeTarget,
} from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';
import { SLIDE_DOCUMENT, type SlideData } from '../_internal-symbols.ts';

const NAME_TIMING = qname('p', 'timing', NS.pml);
const NAME_TN_LST = qname('p', 'tnLst', NS.pml);
const NAME_PAR = qname('p', 'par', NS.pml);
const NAME_C_TN = qname('p', 'cTn', NS.pml);
const NAME_CHILD_TN_LST = qname('p', 'childTnLst', NS.pml);
const NAME_EXT_LST = qname('p', 'extLst', NS.pml);
const ATTR_ID = qname('', 'id', '');

const findSlideTiming = (slide: SlideData): XmlElement | null =>
  firstChildElement(slide[SLIDE_DOCUMENT].root, NAME_TIMING);

// `<p:timing>` sits between `<p:transition>` and `<p:extLst>` in CT_Slide;
// everything else that may precede it is already in place on a parsed slide.
const insertSlideTiming = (slide: SlideData, timing: XmlElement): void => {
  const root = slide[SLIDE_DOCUMENT].root;
  const extLst = firstChildElement(root, NAME_EXT_LST);
  if (extLst === null) root.children.push(timing);
  else root.children.splice(root.children.indexOf(extLst), 0, timing);
};

/** The root time node's `<p:childTnLst>` — where media nodes and sequences live. */
export const rootChildTnLst = (timing: XmlElement): XmlElement | null => {
  const tnLst = firstChildElement(timing, NAME_TN_LST);
  const par = tnLst ? firstChildElement(tnLst, NAME_PAR) : null;
  const cTn = par ? firstChildElement(par, NAME_C_TN) : null;
  return cTn ? firstChildElement(cTn, NAME_CHILD_TN_LST) : null;
};

// Largest numeric `<p:cTn id="N">` anywhere in the tree (0 when none).
export const maxCTnId = (el: XmlElement): number => {
  let max = 0;
  const walk = (e: XmlElement): void => {
    if (e.name.namespaceURI === NS.pml && e.name.localName === 'cTn') {
      const n = Number.parseInt(getAttrValue(e, ATTR_ID) ?? '', 10);
      if (Number.isFinite(n) && n > max) max = n;
    }
    for (const c of e.children) if (c.kind === 'element') walk(c);
  };
  walk(el);
  return max;
};

/** Adds the media time node for shape `spid`. The caller commits the slide. */
export const addMediaTimingNode = (slide: SlideData, kind: MediaFileKind, spid: number): void => {
  const timing = findSlideTiming(slide);
  if (timing === null) {
    // The root node takes cTn id 1, so the first media node is 2.
    insertSlideTiming(slide, buildTimingRoot([buildMediaTimingNode(kind, spid, 2)]));
    return;
  }
  const node = buildMediaTimingNode(kind, spid, maxCTnId(timing) + 1);
  const childTnLst = rootChildTnLst(timing);
  if (childTnLst !== null) {
    childTnLst.children.push(node);
    return;
  }
  // A `<p:timing>` whose root has no `<p:childTnLst>` (e.g. a bare
  // `<p:timing/>`) holds no time nodes worth keeping, so its `<p:tnLst>` is
  // replaced; `<p:bldLst>` / `<p:extLst>` stay where they are.
  const fresh = buildTimingRoot([node]);
  const rest = timing.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'tnLst'),
  );
  timing.children = [...fresh.children, ...rest];
};

/**
 * Removes the media nodes targeting any of `spids`, and the whole `<p:timing>`
 * when that leaves the root without children (an empty `<p:childTnLst>` is
 * schema-invalid and an empty root animates nothing). The caller commits.
 */
export const removeMediaTimingNodes = (slide: SlideData, spids: ReadonlySet<number>): void => {
  const timing = findSlideTiming(slide);
  const childTnLst = timing ? rootChildTnLst(timing) : null;
  if (timing === null || childTnLst === null) return;
  childTnLst.children = childTnLst.children.filter((c) => {
    if (c.kind !== 'element' || !isMediaTimingNode(c)) return true;
    const target = mediaTimingNodeTarget(c);
    return target === null || !spids.has(target);
  });
  if (!childTnLst.children.some((c) => c.kind === 'element')) {
    const root = slide[SLIDE_DOCUMENT].root;
    root.children = root.children.filter((c) => c !== timing);
  }
};

/** Media nodes at the timing root, in document order. */
export const mediaTimingNodes = (timing: XmlElement): XmlElement[] => {
  const childTnLst = rootChildTnLst(timing);
  if (childTnLst === null) return [];
  return childTnLst.children.filter(
    (c): c is XmlElement => c.kind === 'element' && isMediaTimingNode(c),
  );
};
