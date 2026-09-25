// Deck manipulation: add, remove, reorder, duplicate, import, merge.

import {
  type PartName,
  basename,
  emptyRels,
  nextRelId,
  partName,
  relsPartNameFor,
  resolveTarget,
} from '../../internal/opc/index.ts';
import type { OpcPackage } from '../../internal/parts/index.ts';
import { copyPartGraphs, duplicatePartGraph } from '../../internal/parts/duplicate-graph.ts';
import { REL_TYPES, buildSlideFromLayout } from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlDocument,
  type XmlElement,
  allChildElements,
  childElements,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
  serializeXml,
} from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  LAYOUT_PART,
  LAYOUT_PART_NAME,
  type PresentationData,
  SLIDE_DOCUMENT,
  SLIDE_PART_NAME,
  type SlideData,
  type SlideLayoutData,
} from '../_internal-symbols.ts';
import {
  ATTR_ID,
  ATTR_R_ID,
  NAME_CSLD,
  NAME_PRESENTATION,
  NAME_SLD_ID,
  NAME_SLD_ID_LST,
  NAME_SLD_MASTER_ID_LST,
  NAME_SP_TREE,
  PRES_PART_NAME,
  SLD_ID_MAX,
  SLD_ID_MIN,
  SLIDE_CONTENT_TYPE,
  commitSlideData,
  refreshSlideData,
  decode,
  encode,
} from './_helpers.ts';
import {
  findSlideLayoutByType,
  getSlideLayoutType,
  getSlideLayouts,
  getSlideLayoutPlaceholders,
} from './layouts.ts';
import { buildSlideData, getSlides, refreshSlideOrder } from './slide-query.ts';
import { setSlideBody, setSlideTitle } from './embedded.ts';

// ---------------------------------------------------------------------------
// Deck manipulation.

const ensureSldIdLst = (presentationRoot: XmlElement): XmlElement => {
  const existing = firstChildElement(presentationRoot, NAME_SLD_ID_LST);
  if (existing !== null) return existing;
  const fresh = elem(NAME_SLD_ID_LST);
  const masterLst = firstChildElement(presentationRoot, NAME_SLD_MASTER_ID_LST);
  if (masterLst === null) {
    presentationRoot.children.unshift(fresh);
    return fresh;
  }
  const idx = presentationRoot.children.indexOf(masterLst);
  presentationRoot.children.splice(idx + 1, 0, fresh);
  return fresh;
};

const allocateSldId = (sldIdLst: XmlElement): number => {
  let max = SLD_ID_MIN - 1;
  for (const sldId of allChildElements(sldIdLst, NAME_SLD_ID)) {
    const raw = getAttrValue(sldId, ATTR_ID);
    if (raw === null) continue;
    const n = Number.parseInt(raw, 10);
    if (Number.isFinite(n) && n > max) max = n;
  }
  const next = Math.max(SLD_ID_MIN, max + 1);
  if (next > SLD_ID_MAX) {
    throw new Error(`sldId allocator exhausted (next would be ${next}, max ${SLD_ID_MAX})`);
  }
  return next;
};

const allocateSlideN = (pkg: OpcPackage): number => {
  let next = 1;
  for (const p of pkg.parts) {
    const m = p.name.match(/^\/ppt\/slides\/slide(\d+)\.xml$/);
    if (m?.[1] !== undefined) {
      const n = Number.parseInt(m[1], 10);
      if (Number.isFinite(n) && n >= next) next = n + 1;
    }
  }
  return next;
};

const requirePresentationDoc = (pkg: OpcPackage): XmlDocument => {
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) throw new Error('presentation.xml is missing');
  const doc = parseXml(decode(presPart.data));
  if (
    doc.root.name.namespaceURI !== NAME_PRESENTATION.namespaceURI ||
    doc.root.name.localName !== 'presentation'
  ) {
    throw new Error('presentation.xml root is not <p:presentation>');
  }
  return doc;
};

/**
 * Adds a new slide bound to `layout`. Returns the new `SlideData`.
 *
 * Allocates a fresh part name, sldId, and rId; clones layout
 * placeholders into the slide; writes `[Content_Types].xml`, the
 * slide's `.rels`, presentation's `.rels`, and `<p:sldIdLst>`. The
 * deck-cache on `pres` is invalidated so the next `getSlides` call
 * sees the new entry.
 */
/**
 * Convenience over `addSlide` that picks a layout automatically:
 *
 *   1. The layout with `<p:sldLayout type="blank">`, if present.
 *   2. Otherwise, the first available layout (alphabetical by
 *      part name).
 *
 * Throws when the package carries no layouts at all (which would
 * be a structurally-broken deck).
 */
export const addBlankSlide = (pres: PresentationData): SlideData => {
  const blank = findSlideLayoutByType(pres, 'blank');
  if (blank) return addSlide(pres, { layout: blank });
  const layouts = getSlideLayouts(pres);
  if (layouts.length === 0) {
    throw new Error('addBlankSlide: package has no slide layouts to inherit from');
  }
  return addSlide(pres, { layout: layouts[0]! });
};

/**
 * Sugar over `addSlide` + `setSlideTitle` + `setSlideBody` for the
 * "title + body" pattern. Picks the `obj` (Title and Content)
 * layout when present, falling back to the first layout with a
 * body placeholder.
 *
 * Throws if no layout in the package offers a body slot.
 */
export const addContentSlide = (
  pres: PresentationData,
  opts: { title?: string; body?: string },
): SlideData => {
  const objLayout = findSlideLayoutByType(pres, 'obj');
  const layout =
    objLayout ??
    getSlideLayouts(pres).find((l) =>
      getSlideLayoutPlaceholders(l).some((p) => p.type === null || p.type === 'body'),
    );
  if (!layout) {
    throw new Error('addContentSlide: no layout with a body placeholder found');
  }
  const slide = addSlide(pres, { layout });
  if (opts.title !== undefined) setSlideTitle(slide, opts.title);
  if (opts.body !== undefined) setSlideBody(slide, opts.body);
  return slide;
};

/**
 * Sugar over `addSlide` + `setSlideTitle` for the section-divider
 * pattern. Picks `<p:sldLayout type="secHead">` when present (the
 * PowerPoint "Section Header" layout); otherwise falls back to a
 * `title`-typed layout or the first available layout.
 */
export const addSectionHeaderSlide = (pres: PresentationData, title: string): SlideData => {
  const layout =
    findSlideLayoutByType(pres, 'secHead') ??
    findSlideLayoutByType(pres, 'title') ??
    getSlideLayouts(pres)[0];
  if (!layout) {
    throw new Error('addSectionHeaderSlide: package has no slide layouts to inherit from');
  }
  const slide = addSlide(pres, { layout });
  setSlideTitle(slide, title);
  return slide;
};

/**
 * Sugar over `addSlide` + `setSlideTitle` for the common
 * "title slide + set heading" pattern. Picks the `title` layout
 * first, then falls back to the first non-blank layout.
 *
 * Throws when the package carries no layouts at all.
 */
export const addTitleSlide = (pres: PresentationData, title: string): SlideData => {
  const titleLayout =
    findSlideLayoutByType(pres, 'title') ?? findSlideLayoutByType(pres, 'obj') ?? null;
  const layout =
    titleLayout ??
    getSlideLayouts(pres).find((l) => getSlideLayoutType(l) !== 'blank') ??
    getSlideLayouts(pres)[0];
  if (!layout) {
    throw new Error('addTitleSlide: package has no slide layouts to inherit from');
  }
  const slide = addSlide(pres, { layout });
  setSlideTitle(slide, title);
  return slide;
};

export const addSlide = (
  pres: PresentationData,
  options: { layout: SlideLayoutData },
): SlideData => {
  const pkg = pres[INTERNAL_PACKAGE];
  const layout = options.layout;
  const layoutPart = layout[LAYOUT_PART];
  const layoutPartName = layout[LAYOUT_PART_NAME];

  const presDoc = requirePresentationDoc(pkg);
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) throw new Error('presentation.xml is missing');

  const sldIdLst = ensureSldIdLst(presDoc.root);
  const newSldId = allocateSldId(sldIdLst);
  const slideN = allocateSlideN(pkg);
  const newSlidePartName = partName(`/ppt/slides/slide${slideN}.xml`);

  const layoutCsld = firstChildElement(layoutPart.root, NAME_CSLD);
  if (!layoutCsld) throw new Error(`layout ${layoutPartName} missing <p:cSld>`);
  const layoutSpTree = firstChildElement(layoutCsld, NAME_SP_TREE);
  if (!layoutSpTree) throw new Error(`layout ${layoutPartName} missing <p:spTree>`);

  const slideDoc = buildSlideFromLayout(layoutSpTree);
  const slideBytes = encode(serializeXml(slideDoc));
  pkg.addPart(newSlidePartName, SLIDE_CONTENT_TYPE, slideBytes);

  const slideRels = emptyRels();
  slideRels.items.push({
    id: 'rId1',
    type: REL_TYPES.slideLayout,
    target: `../slideLayouts/${basename(layoutPartName)}`,
    targetMode: 'Internal',
  });
  pkg.setRels(newSlidePartName, slideRels);

  const presRels = pkg.getRels(PRES_PART_NAME) ?? emptyRels();
  const newRId = nextRelId(presRels.items.map((r) => r.id));
  presRels.items.push({
    id: newRId,
    type: REL_TYPES.slide,
    target: `slides/slide${slideN}.xml`,
    targetMode: 'Internal',
  });
  pkg.setRels(PRES_PART_NAME, presRels);

  sldIdLst.children.push(
    elem(NAME_SLD_ID, {
      attrs: [attr(ATTR_ID, String(newSldId)), attr(ATTR_R_ID, newRId)],
    }),
  );
  presPart.data = encode(serializeXml(presDoc));

  // Appending does not change existing slide documents. Re-parsing them here
  // made composing a deck quadratic and detached previously returned handles.
  if (pres._slidesCache !== null) {
    const added = buildSlideData(pkg, newSlidePartName, slideBytes);
    pres._slidesCache = [...pres._slidesCache, added];
    return added;
  }
  const slides = getSlides(pres);
  const last = slides[slides.length - 1];
  if (!last) throw new Error('addSlide: post-condition failed; slide not in cache');
  return last;
};

/**
 * Drops relationships in other parts that point at `removed`, and the
 * `<a:hlinkClick>` / `<a:hlinkHover>` elements that carried them.
 *
 * A slide-jump click action stores a `slide` relationship on the *referring*
 * slide. Removing the target leaves that relationship pointing at a deleted
 * part, which PowerPoint rejects and which makes a later `duplicateSlide` of
 * the referring slide fail on the missing dependency.
 */
const dropRelsPointingAtSlide = (pkg: OpcPackage, removed: PartName): void => {
  for (const part of pkg.parts) {
    if (part.name === removed || part.name.endsWith('.rels')) continue;
    const rels = pkg.getRels(part.name);
    if (!rels) continue;

    const dangling = new Set(
      rels.items
        .filter((rel) => {
          if (rel.type !== REL_TYPES.slide || rel.targetMode === 'External') return false;
          const target = rel.target.startsWith('/')
            ? partName(rel.target)
            : resolveTarget(part.name, rel.target);
          return target === removed;
        })
        .map((rel) => rel.id),
    );
    if (dangling.size === 0) continue;

    rels.items = rels.items.filter((rel) => !dangling.has(rel.id));
    pkg.setRels(part.name, rels);

    const doc = parseXml(decode(part.data));
    stripHlinksWithRelId(doc.root, dangling);
    part.data = encode(serializeXml(doc));
  }
};

/** Removes every `<a:hlinkClick>` / `<a:hlinkHover>` whose `r:id` is in `relIds`. */
const stripHlinksWithRelId = (element: XmlElement, relIds: ReadonlySet<string>): void => {
  element.children = element.children.filter((child) => {
    if (child.kind !== 'element') return true;
    const isHlink =
      child.name.namespaceURI === NS.dml &&
      (child.name.localName === 'hlinkClick' || child.name.localName === 'hlinkHover');
    if (!isHlink) return true;
    const rId = getAttrValue(child, ATTR_R_ID);
    return rId === null || !relIds.has(rId);
  });
  for (const child of childElements(element)) stripHlinksWithRelId(child, relIds);
};

/**
 * Removes the given slide from the deck. Removes the `<p:sldId>`, the
 * `presentation.xml.rels` entry, and the slide part + its `.rels` part.
 * Links to this slide from other slides are cleared to prevent accidental
 * retargeting when a slide part name is reused.
 *
 * Media parts are intentionally NOT cleaned up — they may be shared
 * with other slides. The freed `sldId` is NOT reused on subsequent
 * `addSlide` calls (PowerPoint quirk, see plan §Risks).
 */
export const removeSlide = (pres: PresentationData, slide: SlideData): void => {
  const pkg = pres[INTERNAL_PACKAGE];
  const slidePartName = slide[SLIDE_PART_NAME];
  if (slide[INTERNAL_PACKAGE] !== pkg) {
    throw new Error('removeSlide: slide must belong to this presentation');
  }
  if (pkg.getPart(slidePartName) === null) {
    throw new Error(`removeSlide: ${slidePartName} not present in package`);
  }

  const presRels = pkg.getRels(PRES_PART_NAME);
  if (!presRels) throw new Error('presentation.xml has no rels');
  const removedRel = presRels.items.find(
    (r) =>
      r.type === REL_TYPES.slide &&
      r.targetMode !== 'External' &&
      resolveTarget(PRES_PART_NAME, r.target) === slidePartName,
  );
  if (!removedRel) {
    throw new Error(`presentation.xml.rels missing entry for slide ${slidePartName}`);
  }
  presRels.items = presRels.items.filter((r) => r.id !== removedRel.id);
  pkg.setRels(PRES_PART_NAME, presRels);

  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) throw new Error('presentation.xml missing');
  const presDoc = parseXml(decode(presPart.data));
  const sldIdLst = firstChildElement(presDoc.root, NAME_SLD_ID_LST);
  if (sldIdLst !== null) {
    sldIdLst.children = sldIdLst.children.filter((c) => {
      if (c.kind !== 'element') return true;
      if (c.name.namespaceURI !== NS.pml || c.name.localName !== 'sldId') return true;
      return getAttrValue(c, ATTR_R_ID) !== removedRel.id;
    });
  }
  presPart.data = encode(serializeXml(presDoc));

  // Remove inbound slide links before their part name or relationship id can be
  // reused. Edit the live document so retained shape/cell handles stay in sync.
  for (const source of getSlides(pres)) {
    if (source[SLIDE_PART_NAME] === slidePartName) continue;
    const rels = pkg.getRels(source[SLIDE_PART_NAME]);
    if (!rels) continue;
    const removedIds = new Set(
      rels.items
        .filter(
          (rel) =>
            rel.type === REL_TYPES.slide &&
            rel.targetMode !== 'External' &&
            resolveTarget(source[SLIDE_PART_NAME], rel.target) === slidePartName,
        )
        .map((rel) => rel.id),
    );
    if (!removedIds.size) continue;
    const removeLinks = (node: XmlElement): void => {
      node.children = node.children.filter((child) => {
        if (child.kind !== 'element') return true;
        if (
          child.name.namespaceURI === NS.dml &&
          ['hlinkClick', 'hlinkMouseOver', 'hlinkHover'].includes(child.name.localName) &&
          removedIds.has(getAttrValue(child, ATTR_R_ID) ?? '')
        )
          return false;
        removeLinks(child);
        return true;
      });
    };
    removeLinks(source[SLIDE_DOCUMENT].root);
    rels.items = rels.items.filter((rel) => !removedIds.has(rel.id));
    pkg.setRels(source[SLIDE_PART_NAME], rels);
    commitSlideData(source);
    refreshSlideData(source);
  }

  pkg.removePart(relsPartNameFor(slidePartName));
  pkg.removePart(slidePartName);
  dropRelsPointingAtSlide(pkg, slidePartName);
  // Rebuild from the deck's own order rather than dropping the cache: the
  // surviving slides keep the handles the caller (and the editor) still holds.
  refreshSlideOrder(pres);
};

/**
 * Reorders every slide in the deck via a custom comparator. The
 * comparator is invoked with two `SlideData` handles and returns the
 * usual `Array.prototype.sort` ordering (-1 / 0 / 1).
 *
 *   sortSlides(pres, (a, b) => getSlideTitle(a)?.localeCompare(getSlideTitle(b) ?? '') ?? 0);
 *
 * Internally walks `<p:sldIdLst>` and re-emits its `<p:sldId>` children
 * in the new order. Slide parts and rels are untouched — only the
 * order in which PowerPoint plays them changes.
 */
export const sortSlides = (
  pres: PresentationData,
  compareFn: (a: SlideData, b: SlideData) => number,
): void => {
  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) throw new Error('presentation.xml missing');
  const doc = parseXml(decode(presPart.data));
  const sldIdLst = firstChildElement(doc.root, NAME_SLD_ID_LST);
  if (!sldIdLst) return; // nothing to reorder

  const slides = getSlides(pres);
  const presRels = pkg.getRels(PRES_PART_NAME);
  if (!presRels) return;

  // Build a map from rId → SlideData and from rId → its <p:sldId> element.
  const slideByRId = new Map<string, SlideData>();
  for (const slide of slides) {
    const rel = presRels.items.find(
      (r) =>
        r.type === REL_TYPES.slide && r.target === `slides/${basename(slide[SLIDE_PART_NAME])}`,
    );
    if (rel) slideByRId.set(rel.id, slide);
  }
  const sldIdElements = sldIdLst.children.filter(
    (c): c is XmlElement =>
      c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'sldId',
  );
  const sortedSlides = [...slides].sort(compareFn);
  const newOrder: XmlElement[] = [];
  for (const slide of sortedSlides) {
    let matchedRId: string | undefined;
    for (const [rId, s] of slideByRId.entries()) {
      if (s === slide) {
        matchedRId = rId;
        break;
      }
    }
    if (matchedRId === undefined) continue;
    const el = sldIdElements.find((e) => getAttrValue(e, ATTR_R_ID) === matchedRId);
    if (el) newOrder.push(el);
  }

  // Replace the children, preserving any non-sldId children (whitespace
  // or comments — unlikely but defensive).
  const nonSldIdChildren = sldIdLst.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'sldId'),
  );
  sldIdLst.children = [...nonSldIdChildren, ...newOrder];
  presPart.data = encode(serializeXml(doc));
  refreshSlideOrder(pres);
};

/**
 * Reverses the slide order across the whole deck. Built on
 * `sortSlides` for predictable rels behavior.
 */
export const reverseSlides = (pres: PresentationData): void => {
  const indexBy = new Map<SlideData, number>();
  for (const [i, slide] of getSlides(pres).entries()) indexBy.set(slide, i);
  sortSlides(pres, (a, b) => (indexBy.get(b) ?? 0) - (indexBy.get(a) ?? 0));
};

/**
 * Swaps the positions of the slides at `indexA` and `indexB`.
 * No-op when the indices are equal. Throws on out-of-range indices.
 * Implemented on top of `moveSlide` for predictable rels behavior.
 */
export const swapSlides = (pres: PresentationData, indexA: number, indexB: number): void => {
  if (indexA === indexB) return;
  const slides = getSlides(pres);
  const a = slides[indexA];
  const b = slides[indexB];
  if (!a) throw new RangeError(`swapSlides: indexA ${indexA} out of range (have ${slides.length})`);
  if (!b) throw new RangeError(`swapSlides: indexB ${indexB} out of range (have ${slides.length})`);
  // Move the lower-index slide to the higher index first so the
  // remaining slide stays at its original index.
  const [lo, hi] = indexA < indexB ? [indexA, indexB] : [indexB, indexA];
  moveSlide(pres, slides[lo]!, hi);
  // After the first move, the slide originally at hi is now at hi-1.
  const refreshed = getSlides(pres);
  moveSlide(pres, refreshed[hi - 1]!, lo);
};

/**
 * Reorders a slide. The slide's identity (part, rels, sldId) is
 * unchanged — only `<p:sldIdLst>`'s child order changes.
 */
export const moveSlide = (pres: PresentationData, slide: SlideData, toIndex: number): void => {
  const pkg = pres[INTERNAL_PACKAGE];
  const slideRelTarget = `slides/${basename(slide[SLIDE_PART_NAME])}`;
  const presRels = pkg.getRels(PRES_PART_NAME);
  if (!presRels) throw new Error('presentation.xml has no rels');
  const slideRel = presRels.items.find(
    (r) => r.type === REL_TYPES.slide && r.target === slideRelTarget,
  );
  if (!slideRel) throw new Error(`moveSlide: slide ${slide[SLIDE_PART_NAME]} has no rel`);

  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) throw new Error('presentation.xml missing');
  const presDoc = parseXml(decode(presPart.data));
  const sldIdLst = firstChildElement(presDoc.root, NAME_SLD_ID_LST);
  if (!sldIdLst) throw new Error('presentation.xml has no <p:sldIdLst>');

  const sldIdElements = sldIdLst.children.filter(
    (c): c is XmlElement =>
      c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'sldId',
  );
  const target = sldIdElements.find((e) => getAttrValue(e, ATTR_R_ID) === slideRel.id);
  if (!target) throw new Error(`moveSlide: <p:sldId> for ${slideRel.id} not found`);

  const remaining = sldIdLst.children.filter((c) => c !== target);
  const remainingSldIds = remaining.filter(
    (c): c is XmlElement =>
      c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'sldId',
  );
  const clamped = Math.max(0, Math.min(toIndex, remainingSldIds.length));
  if (clamped === remainingSldIds.length) {
    remaining.push(target);
  } else {
    const before = remainingSldIds[clamped];
    const insertAt = before === undefined ? remaining.length : remaining.indexOf(before);
    remaining.splice(insertAt, 0, target);
  }
  sldIdLst.children = remaining;
  presPart.data = encode(serializeXml(presDoc));
  refreshSlideOrder(pres);
};

/**
 * Duplicates a slide. Returns the new `SlideData` appended to deck order.
 *
 * Charts, embedded workbooks, notes and other owned dependencies are copied.
 * Layouts, masters, themes, media and links to other slides remain shared.
 * Unknown dependency bodies are retained without interpreting their content.
 */
export const duplicateSlide = (pres: PresentationData, slide: SlideData): SlideData => {
  const pkg = pres[INTERNAL_PACKAGE];
  if (slide[INTERNAL_PACKAGE] !== pkg)
    throw new Error('duplicateSlide: slide belongs to another presentation');
  const sourcePartName = slide[SLIDE_PART_NAME];
  const sourcePart = pkg.getPart(sourcePartName);
  if (!sourcePart) throw new Error(`duplicateSlide: source ${sourcePartName} not found`);

  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) throw new Error('presentation.xml missing');
  const presDoc = parseXml(decode(presPart.data));
  const sldIdLst = ensureSldIdLst(presDoc.root);
  const newSldId = allocateSldId(sldIdLst);

  const slideN = allocateSlideN(pkg);
  const newSlidePartName = partName(`/ppt/slides/slide${slideN}.xml`);
  duplicatePartGraph(
    pkg,
    sourcePartName,
    newSlidePartName,
    new Set([
      REL_TYPES.slideLayout,
      REL_TYPES.slideMaster,
      REL_TYPES.notesMaster,
      REL_TYPES.handoutMaster,
      REL_TYPES.theme,
      REL_TYPES.slide,
      REL_TYPES.image,
      REL_TYPES.media,
      REL_TYPES.video,
      REL_TYPES.audio,
    ]),
  );

  const presRels = pkg.getRels(PRES_PART_NAME) ?? emptyRels();
  const newRId = nextRelId(presRels.items.map((r) => r.id));
  presRels.items.push({
    id: newRId,
    type: REL_TYPES.slide,
    target: `slides/slide${slideN}.xml`,
    targetMode: 'Internal',
  });
  pkg.setRels(PRES_PART_NAME, presRels);

  sldIdLst.children.push(
    elem(NAME_SLD_ID, {
      attrs: [attr(ATTR_ID, String(newSldId)), attr(ATTR_R_ID, newRId)],
    }),
  );
  presPart.data = encode(serializeXml(presDoc));

  if (pres._slidesCache !== null) {
    const added = buildSlideData(pkg, newSlidePartName, pkg.getPart(newSlidePartName)!.data);
    pres._slidesCache = [...pres._slidesCache, added];
    return added;
  }
  const slides = getSlides(pres);
  const dup = slides[slides.length - 1];
  if (!dup) throw new Error('duplicateSlide: post-condition failed');
  return dup;
};

/**
 * Convenience over `addSlide` + `moveSlide`. Inserts the new slide
 * at the given 0-based index (clamped to `[0, getSlides(pres).length]`).
 */
export const addSlideAt = (
  pres: PresentationData,
  atIndex: number,
  options: { layout: SlideLayoutData },
): SlideData => {
  const slide = addSlide(pres, options);
  moveSlide(pres, slide, atIndex);
  const slides = getSlides(pres);
  const clamped = Math.max(0, Math.min(atIndex, slides.length - 1));
  return slides[clamped]!;
};

/**
 * Convenience over `duplicateSlide` + `moveSlide`. Duplicates `slide`
 * and inserts the duplicate at `atIndex` instead of at the end.
 */
export const duplicateSlideAt = (
  pres: PresentationData,
  atIndex: number,
  slide: SlideData,
): SlideData => {
  const dup = duplicateSlide(pres, slide);
  moveSlide(pres, dup, atIndex);
  const slides = getSlides(pres);
  const clamped = Math.max(0, Math.min(atIndex, slides.length - 1));
  return slides[clamped]!;
};

/**
 * Imports a slide into `targetPres`. When `targetLayout` is supplied, rebinds
 * the slide to it. Otherwise preserves the source layout, master and theme.
 * Copies the complete relationship graph, including charts, workbooks, notes,
 * media and unknown extension parts. Shared dependencies and cycles retain
 * their relationships; external hyperlinks are preserved. Missing dependencies
 * fail before writing any imported parts.
 * Returns the new slide appended to the target presentation.
 */
export const importSlide = (
  targetPres: PresentationData,
  sourceSlide: SlideData,
  targetLayout?: SlideLayoutData,
): SlideData => {
  const sourcePkg = sourceSlide[INTERNAL_PACKAGE];
  const sourcePartName = sourceSlide[SLIDE_PART_NAME];
  const sourcePart = sourcePkg.getPart(sourcePartName);
  if (!sourcePart) throw new Error(`importSlide: source ${sourcePartName} not found`);
  const sourceRels = sourcePkg.getRels(sourcePartName);

  const targetPkg = targetPres[INTERNAL_PACKAGE];
  const presPart = targetPkg.getPart(PRES_PART_NAME);
  if (!presPart) throw new Error('presentation.xml missing in target');
  const presDoc = parseXml(decode(presPart.data));
  const sldIdLst = ensureSldIdLst(presDoc.root);
  const newSldId = allocateSldId(sldIdLst);

  const slideN = allocateSlideN(targetPkg);
  const newSlidePartName = partName(`/ppt/slides/slide${slideN}.xml`);

  const layoutPartName = targetLayout?.[LAYOUT_PART_NAME];
  if (layoutPartName && targetPkg.getPart(layoutPartName) === null) {
    throw new Error(`importSlide: layout ${layoutPartName} not in target package`);
  }
  const layoutRel = sourceRels?.items.find((rel) => rel.type === REL_TYPES.slideLayout);
  const mappedLayouts = new Map<PartName, PartName>();
  if (layoutRel && layoutPartName)
    mappedLayouts.set(resolveTarget(sourcePartName, layoutRel.target), layoutPartName);
  const copies = copyPartGraphs(
    sourcePkg,
    targetPkg,
    new Map([[sourcePartName, newSlidePartName]]),
    new Set(),
    mappedLayouts,
  );
  if (!layoutRel && layoutPartName) {
    const rels = targetPkg.getRels(newSlidePartName) ?? emptyRels();
    rels.items.push({
      id: nextRelId(rels.items.map((rel) => rel.id)),
      type: REL_TYPES.slideLayout,
      target: layoutPartName,
      targetMode: 'Internal',
    });
    targetPkg.setRels(newSlidePartName, rels);
  }

  // presentation → slide rel + sldIdLst entry.
  const presRels = targetPkg.getRels(PRES_PART_NAME) ?? emptyRels();
  // Imported masters must be discoverable from presentation.xml as well as layouts.
  const masterName = qname('p', 'sldMasterId', NS.pml);
  const masters = [...copies.values()].filter((name) =>
    targetPkg.getPart(name)?.contentType.endsWith('slideMaster+xml'),
  );
  if (masters.length) {
    let masterList = firstChildElement(presDoc.root, NAME_SLD_MASTER_ID_LST);
    if (!masterList) {
      masterList = elem(NAME_SLD_MASTER_ID_LST);
      presDoc.root.children.unshift(masterList);
    }
    const used = new Set(
      allChildElements(masterList, masterName).map((item) => Number(getAttrValue(item, ATTR_ID))),
    );
    let id = 2147483648;
    for (const master of masters) {
      while (used.has(id)) id++;
      used.add(id);
      const relId = nextRelId(presRels.items.map((rel) => rel.id));
      presRels.items.push({
        id: relId,
        type: REL_TYPES.slideMaster,
        target: master,
        targetMode: 'Internal',
      });
      masterList.children.push(
        elem(masterName, { attrs: [attr(ATTR_ID, String(id)), attr(ATTR_R_ID, relId)] }),
      );
    }
  }
  const newRId = nextRelId(presRels.items.map((r) => r.id));
  presRels.items.push({
    id: newRId,
    type: REL_TYPES.slide,
    target: `slides/slide${slideN}.xml`,
    targetMode: 'Internal',
  });
  targetPkg.setRels(PRES_PART_NAME, presRels);

  sldIdLst.children.push(
    elem(NAME_SLD_ID, {
      attrs: [attr(ATTR_ID, String(newSldId)), attr(ATTR_R_ID, newRId)],
    }),
  );
  presPart.data = encode(serializeXml(presDoc));

  refreshSlideOrder(targetPres);
  const slides = getSlides(targetPres);
  const last = slides[slides.length - 1];
  if (!last) throw new Error('importSlide: post-condition failed');
  return last;
};

/**
 * Appends every slide from `sourcePres` into `targetPres`, in source
 * order. Built on top of `importSlide`: dependencies are preserved, and the slide's
 * layout is rebound to `targetLayout` on the target side.
 *
 * `targetLayout` can be a single layout used for every imported
 * slide (common), or a function called once per source slide for
 * per-slide layout selection.
 *
 * Returns the imported slides in target order.
 */
export const mergePresentations = (
  targetPres: PresentationData,
  sourcePres: PresentationData,
  targetLayout: SlideLayoutData | ((sourceSlide: SlideData, index: number) => SlideLayoutData),
): ReadonlyArray<SlideData> => {
  const sourceSlides = getSlides(sourcePres);
  const out: SlideData[] = [];
  const resolveLayout =
    typeof targetLayout === 'function' ? targetLayout : (): SlideLayoutData => targetLayout;
  for (let i = 0; i < sourceSlides.length; i++) {
    const src = sourceSlides[i]!;
    const layout = resolveLayout(src, i);
    out.push(importSlide(targetPres, src, layout));
  }
  return out;
};
