// Tree-shakeable free-function entry points.
//
// Every operation in this module is a standalone export that operates on
// the opaque `PresentationData` / `SlideData` interfaces (shared with the
// class-based API via `_internal-symbols.ts`). Crucially, none of these
// functions references any class — when a consumer imports only what they
// need from this module, modern bundlers drop the class definitions in
// `presentation.ts` / `slide.ts` / etc. entirely.
//
// The class-based API in those files is preserved as a legacy facade
// (re-exported from `./index.ts` as `Presentation`, `Slide`, etc.) so
// existing tests and downstream consumers continue to work. As class
// methods migrate to live exclusively here, the class definitions will
// shrink and eventually disappear.

import {
  type Position,
  type Size,
  readFlip,
  readPosition,
  readRotation,
  readSize,
  replaceTokensInTree,
} from '../internal/drawingml/index.ts';
import {
  basename,
  emptyRels,
  nextRelId,
  type PartName,
  partName,
  relsPartNameFor,
  resolveTarget,
} from '../internal/opc/index.ts';
import { OpcPackage } from '../internal/parts/index.ts';
import {
  REL_TYPES,
  type ShapeKind,
  buildSlideFromLayout,
  readPresentationPart,
  readSlideLayoutPart,
  readSlidePart,
  slideText,
} from '../internal/presentationml/index.ts';
import {
  NS,
  type XmlDocument,
  type XmlElement,
  allChildElements,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
  serializeXml,
} from '../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  LAYOUT_PART,
  LAYOUT_PART_NAME,
  type PresentationData,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  SLIDE_DOCUMENT,
  SLIDE_PART,
  SLIDE_PART_NAME,
  SLIDE_SHAPES,
  type SlideData,
  type SlideLayoutData,
  type SlideShapeData,
} from './_internal-symbols.ts';

const TEXT_DECODER = new TextDecoder();
const TEXT_ENCODER = new TextEncoder();
const decode = (b: Uint8Array): string => TEXT_DECODER.decode(b);
const encode = (s: string): Uint8Array => TEXT_ENCODER.encode(s);

const SLIDE_LAYOUT_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml';
const SLIDE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const PRES_PART_NAME = partName('/ppt/presentation.xml');

const NAME_PRESENTATION = qname('p', 'presentation', NS.pml);
const NAME_SLD_MASTER_ID_LST = qname('p', 'sldMasterIdLst', NS.pml);
const NAME_SLD_ID_LST = qname('p', 'sldIdLst', NS.pml);
const NAME_SLD_ID = qname('p', 'sldId', NS.pml);
const NAME_CSLD = qname('p', 'cSld', NS.pml);
const NAME_SP_TREE = qname('p', 'spTree', NS.pml);
const ATTR_ID = qname('', 'id', '');
const ATTR_R_ID = qname('r', 'id', NS.officeDocRels);

// PowerPoint accepts sldIds in [256, 2³¹−1024]. See plan §Risks.
const SLD_ID_MIN = 256;
const SLD_ID_MAX = 2147482623;

/**
 * Anything that can be turned into a `Uint8Array` of PPTX bytes:
 *
 *   - `Uint8Array` — used as-is.
 *   - `ArrayBuffer` — wrapped without copying.
 *   - `Blob` or `File` — read via `arrayBuffer()`.
 */
export type PresentationInput = Uint8Array | ArrayBuffer | Blob;

const normalize = async (input: PresentationInput): Promise<Uint8Array> => {
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  if (typeof Blob !== 'undefined' && input instanceof Blob) {
    return new Uint8Array(await input.arrayBuffer());
  }
  throw new TypeError('loadPresentation: expected Uint8Array | ArrayBuffer | Blob');
};

/**
 * Loads an existing `.pptx` and returns a `PresentationData` value.
 */
export const loadPresentation = async (input: PresentationInput): Promise<PresentationData> => {
  const bytes = await normalize(input);
  const pkg = OpcPackage.load(bytes);
  return { [INTERNAL_PACKAGE]: pkg, _slidesCache: null };
};

/**
 * Creates a fresh, empty `PresentationData`. The result is NOT yet a
 * valid PPTX — it carries only the OPC defaults.
 */
export const createPresentation = (): PresentationData => {
  const pkg = OpcPackage.empty();
  return { [INTERNAL_PACKAGE]: pkg, _slidesCache: null };
};

/**
 * Serializes a presentation back to PPTX bytes.
 */
export const savePresentation = (pres: PresentationData): Promise<Uint8Array> => {
  return Promise.resolve(pres[INTERNAL_PACKAGE].save());
};

// ---------------------------------------------------------------------------
// Slide layouts.

/**
 * Enumerates every slide layout in the package.
 */
export const getSlideLayouts = (pres: PresentationData): ReadonlyArray<SlideLayoutData> => {
  const pkg = pres[INTERNAL_PACKAGE];
  const out: SlideLayoutData[] = [];
  for (const part of pkg.parts) {
    if (part.contentType !== SLIDE_LAYOUT_CONTENT_TYPE) continue;
    const root = parseXml(decode(part.data)).root;
    out.push({
      [LAYOUT_PART_NAME]: part.name,
      [LAYOUT_PART]: readSlideLayoutPart(root),
    });
  }
  return out;
};

// ---------------------------------------------------------------------------
// SlideData factory + cached enumeration.

const buildSlideData = (
  pkg: OpcPackage,
  partNameValue: PartName,
  bytes: Uint8Array,
): SlideData => {
  const doc = parseXml(decode(bytes));
  const part = readSlidePart(doc.root);
  const shapes: SlideShapeData[] = [];
  const slide: SlideData = {
    [INTERNAL_PACKAGE]: pkg,
    [SLIDE_PART_NAME]: partNameValue,
    [SLIDE_DOCUMENT]: doc,
    [SLIDE_PART]: part,
    [SLIDE_SHAPES]: shapes,
  };
  for (const snap of part.shapes) {
    shapes.push({
      [SHAPE_SLIDE]: slide,
      [SHAPE_ELEMENT]: snap.element,
      [SHAPE_SNAPSHOT]: snap,
    });
  }
  return slide;
};

/**
 * Enumerates slides in presentation order. Returns plain `SlideData`
 * values — opaque handles whose internal symbols are shared with the
 * class API so either representation can be passed to slide-level
 * functions.
 *
 * Throws if any referenced slide part is missing — a structurally
 * invalid PPTX cannot honor the L1 contract.
 */
export const getSlides = (pres: PresentationData): ReadonlyArray<SlideData> => {
  const cached = pres._slidesCache;
  if (cached !== null) return cached as ReadonlyArray<SlideData>;

  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (presPart === null) {
    const empty: SlideData[] = [];
    pres._slidesCache = empty;
    return empty;
  }
  const presRels = pkg.getRels(PRES_PART_NAME);
  if (presRels === null) {
    const empty: SlideData[] = [];
    pres._slidesCache = empty;
    return empty;
  }

  const presRoot = parseXml(decode(presPart.data)).root;
  const presModel = readPresentationPart(presRoot);

  const out: SlideData[] = [];
  for (const sld of presModel.slides) {
    const rel = presRels.items.find((r) => r.id === sld.rId);
    if (!rel) throw new Error(`presentation.xml.rels missing entry for ${sld.rId}`);
    const target = rel.target;
    const slideName = partName(target.startsWith('/') ? target : `/ppt/${target}`);
    const slidePart = pkg.getPart(slideName);
    if (slidePart === null) throw new Error(`slide part ${slideName} not found`);
    out.push(buildSlideData(pkg, slideName, slidePart.data));
  }
  pres._slidesCache = out;
  return out;
};

/**
 * Concatenated visible text of a slide. Convenience wrapper that walks
 * the slide's shape tree without instantiating any class.
 */
export const getSlideText = (slide: SlideData): string => slideText(slide[SLIDE_PART]);

/**
 * Replaces `{{key}}` tokens on every slide. Returns the total number of
 * substitutions performed.
 */
export const replaceTokensInPresentation = (
  pres: PresentationData,
  tokens: Record<string, string>,
): number => {
  const pkg = pres[INTERNAL_PACKAGE];
  let count = 0;
  for (const part of pkg.parts) {
    if (part.contentType !== SLIDE_CONTENT_TYPE) continue;
    const doc = parseXml(decode(part.data));
    const n = replaceTokensInTree(doc.root, tokens);
    if (n > 0) {
      part.data = encode(serializeXml(doc));
      count += n;
    }
  }
  pres._slidesCache = null;
  return count;
};

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
 * Mirrors `Presentation.addSlide`: allocates a fresh part name, sldId,
 * and rId; clones layout placeholders into the slide; writes
 * `[Content_Types].xml`, the slide's `.rels`, presentation's `.rels`,
 * and `<p:sldIdLst>`. The deck-cache on `pres` is invalidated so the
 * next `getSlides` call sees the new entry.
 */
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

  pres._slidesCache = null;
  const slides = getSlides(pres);
  const last = slides[slides.length - 1];
  if (!last) throw new Error('addSlide: post-condition failed; slide not in cache');
  return last;
};

/**
 * Removes the given slide from the deck. Removes the `<p:sldId>`, the
 * `presentation.xml.rels` entry, and the slide part + its `.rels` part.
 *
 * Media parts are intentionally NOT cleaned up — they may be shared
 * with other slides. The freed `sldId` is NOT reused on subsequent
 * `addSlide` calls (PowerPoint quirk, see plan §Risks).
 */
export const removeSlide = (pres: PresentationData, slide: SlideData): void => {
  const pkg = pres[INTERNAL_PACKAGE];
  const slidePartName = slide[SLIDE_PART_NAME];
  if (pkg.getPart(slidePartName) === null) {
    throw new Error(`removeSlide: ${slidePartName} not present in package`);
  }

  const presRels = pkg.getRels(PRES_PART_NAME);
  if (!presRels) throw new Error('presentation.xml has no rels');
  const slideTargetRel = `slides/${basename(slidePartName)}`;
  const removedRel = presRels.items.find(
    (r) => r.type === REL_TYPES.slide && r.target === slideTargetRel,
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

  pkg.removePart(relsPartNameFor(slidePartName));
  pkg.removePart(slidePartName);
  pres._slidesCache = null;
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
  pres._slidesCache = null;
};

/**
 * Duplicates a slide. Returns the new `SlideData` appended to deck order.
 *
 * Part bytes and rels are cloned verbatim; media parts are NOT copied —
 * the duplicate shares the original's media references (PowerPoint
 * does the same).
 */
export const duplicateSlide = (pres: PresentationData, slide: SlideData): SlideData => {
  const pkg = pres[INTERNAL_PACKAGE];
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
  pkg.addPart(newSlidePartName, sourcePart.contentType, new Uint8Array(sourcePart.data));

  const sourceRels = pkg.getRels(sourcePartName);
  if (sourceRels !== null) {
    pkg.setRels(newSlidePartName, { items: sourceRels.items.map((r) => ({ ...r })) });
  }

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

  pres._slidesCache = null;
  const slides = getSlides(pres);
  const dup = slides[slides.length - 1];
  if (!dup) throw new Error('duplicateSlide: post-condition failed');
  return dup;
};

// ---------------------------------------------------------------------------
// Slide-level reads.

/**
 * Shapes on a slide, in document order with group children flattened.
 */
export const getSlideShapes = (slide: SlideData): ReadonlyArray<SlideShapeData> =>
  slide[SLIDE_SHAPES];

/**
 * The slide layout this slide is bound to, or `null` if the slide has
 * no layout relationship.
 */
export const getSlideLayout = (slide: SlideData): SlideLayoutData | null => {
  const pkg = slide[INTERNAL_PACKAGE];
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (rels === null) return null;
  const layoutRel = rels.items.find((r) => r.type === REL_TYPES.slideLayout);
  if (!layoutRel) return null;
  const layoutName = layoutRel.target.startsWith('/')
    ? partName(layoutRel.target)
    : resolveTarget(slide[SLIDE_PART_NAME], layoutRel.target);
  const layoutPart = pkg.getPart(layoutName);
  if (layoutPart === null) return null;
  const root = parseXml(decode(layoutPart.data)).root;
  return {
    [LAYOUT_PART_NAME]: layoutName,
    [LAYOUT_PART]: readSlideLayoutPart(root),
  };
};

/**
 * Returns the first placeholder shape with the given `type` (or `null`
 * if no match). Shapes whose `<p:ph>` omits an explicit type default to
 * `'body'` per ECMA-376 §19.7.10.
 */
export const findSlidePlaceholder = (
  slide: SlideData,
  type: string,
): SlideShapeData | null => {
  for (const shape of slide[SLIDE_SHAPES]) {
    const snap = shape[SHAPE_SNAPSHOT];
    if (snap.placeholderType === type) return shape;
    if (type === 'body' && snap.placeholderType === null && snap.placeholderIdx !== null) {
      return shape;
    }
  }
  return null;
};

/**
 * Replaces `{{key}}` tokens in every text-bearing shape on this slide.
 * Returns the number of substitutions performed.
 *
 * Tokens must fit within a single text run (see `replaceTokensInTree`
 * in `drawingml/`). Cross-run replacements aren't supported — use
 * `findSlidePlaceholder` + a setText path when PowerPoint has
 * fragmented the run sequence.
 */
export const replaceTokensInSlide = (
  slide: SlideData,
  tokens: Record<string, string>,
): number => {
  const n = replaceTokensInTree(slide[SLIDE_DOCUMENT].root, tokens);
  if (n > 0) {
    commitSlideData(slide);
    refreshSlideData(slide);
  }
  return n;
};

// ---------------------------------------------------------------------------
// SlideShape-level reads.

export const getShapeKind = (shape: SlideShapeData): ShapeKind =>
  shape[SHAPE_SNAPSHOT].kind;

export const getShapeId = (shape: SlideShapeData): number => shape[SHAPE_SNAPSHOT].id;

export const getShapeName = (shape: SlideShapeData): string =>
  shape[SHAPE_SNAPSHOT].name;

export const getShapePlaceholderType = (shape: SlideShapeData): string | null =>
  shape[SHAPE_SNAPSHOT].placeholderType;

export const getShapePlaceholderIdx = (shape: SlideShapeData): number | null =>
  shape[SHAPE_SNAPSHOT].placeholderIdx;

export const getShapeText = (shape: SlideShapeData): string =>
  shape[SHAPE_SNAPSHOT].text;

export const getShapePosition = (shape: SlideShapeData): Position | null =>
  readPosition(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind);

export const getShapeSize = (shape: SlideShapeData): Size | null =>
  readSize(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind);

export const getShapeRotation = (shape: SlideShapeData): number =>
  readRotation(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind);

export const getShapeFlip = (
  shape: SlideShapeData,
): { horizontal: boolean; vertical: boolean } | null =>
  readFlip(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind);

// ---------------------------------------------------------------------------
// @internal — used by mutation functions to write SlideData state back
// into the package and rebuild the typed view. Free functions, no class
// dependency.

const commitSlideData = (slide: SlideData): void => {
  const xml = serializeXml(slide[SLIDE_DOCUMENT]);
  const part = slide[INTERNAL_PACKAGE].getPart(slide[SLIDE_PART_NAME]);
  if (!part) throw new Error(`slide part missing: ${slide[SLIDE_PART_NAME]}`);
  part.data = encode(xml);
};

const refreshSlideData = (slide: SlideData): void => {
  const fresh = readSlidePart(slide[SLIDE_DOCUMENT].root);
  slide[SLIDE_PART] = fresh;
  const shapes = slide[SLIDE_SHAPES];
  for (let i = 0; i < shapes.length; i++) {
    const next = fresh.shapes[i];
    const existing = shapes[i];
    if (!next || !existing) continue;
    existing[SHAPE_ELEMENT] = next.element;
    existing[SHAPE_SNAPSHOT] = next;
  }
};
