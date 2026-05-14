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
  type BulletStyle,
  type GradientFillOptions,
  type ParagraphAlignment,
  type Position,
  type Size,
  type StrokeOptions,
  type TextFormat,
  applyAlignmentToAllParagraphs,
  applyBulletToAllParagraphs,
  applyFormatToAllRuns,
  applyHyperlinkToAllRuns,
  applyRunFormat as applyRunFormatInternal,
  clearFill as clearFillImpl,
  clearStroke as clearStrokeImpl,
  getPictureEmbedRId,
  readFlip,
  readPosition,
  readRotation,
  readSize,
  replaceTokensInTree,
  setFlip as writeFlip,
  setGradientFill,
  setNoFill as setNoFillImpl,
  setNoStroke as setNoStrokeImpl,
  setPosition as writePosition,
  setRotation as writeRotation,
  setSize as writeSize,
  setSolidFill,
  setSolidStroke,
  setTextBody,
} from '../internal/drawingml/index.ts';
import type { Emu } from './units.ts';
import {
  basename,
  contentTypeForFormat,
  detectImageFormat,
  emptyRels,
  extensionForFormat,
  type ImageFormat,
  nextRelId,
  type PartName,
  partName,
  relsPartNameFor,
  resolveTarget,
} from '../internal/opc/index.ts';
import { OpcPackage } from '../internal/parts/index.ts';
import {
  REL_TYPES,
  type AnimationEffect,
  type AnimationOptions,
  type CommentAuthor,
  type CommentPosition,
  type PresetShape,
  type ShapeKind,
  type SlideComment,
  type SlideLayoutType,
  type TransitionOptions,
  buildCommentAuthorListDoc,
  buildCommentListDoc,
  buildConnector,
  buildEmptyNotesSlide,
  buildPicture,
  buildShape,
  buildSingleEffectTiming,
  buildSlideFromLayout,
  buildTable,
  buildTextBox,
  buildTransition,
  readCommentAuthorList,
  readCommentList,
  readPresentationPart,
  readSlideLayoutPart,
  readSlidePart,
  slideText,
} from '../internal/presentationml/index.ts';
import {
  type IssueSeverity,
  type ValidationIssue,
  validatePresentationPackage,
} from '../internal/validator/index.ts';
import {
  type ChartKind,
  type ChartSeries,
  type ChartSpec,
  buildChartSpaceDoc,
  buildEmbeddedXlsx,
} from '../internal/chartml/index.ts';
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
  text as textNode,
} from '../internal/xml/index.ts';
import {
  COMMENT_SLIDE,
  COMMENT_SNAPSHOT,
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
  type SlideCommentData,
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

/** PowerPoint's user-visible layout name. */
export const getSlideLayoutName = (layout: SlideLayoutData): string =>
  layout[LAYOUT_PART].name;

/**
 * Finds the first slide layout whose user-visible name matches `name`,
 * or `null` if none does. Convenience over `getSlideLayouts(...).find(...)`.
 */
export const findSlideLayout = (
  pres: PresentationData,
  name: string,
): SlideLayoutData | null => {
  for (const layout of getSlideLayouts(pres)) {
    if (layout[LAYOUT_PART].name === name) return layout;
  }
  return null;
};

/**
 * Layout type token, when present (`title`, `obj`, `twoObj`, ...).
 * `null` when omitted — the spec default for that case is `cust`.
 */
export const getSlideLayoutType = (
  layout: SlideLayoutData,
): SlideLayoutType | string | null => layout[LAYOUT_PART].layoutType;

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
 * First shape on the slide whose `cNvPr@name` equals `name`, or `null`
 * if none. Use the multi-match variant when more than one shape can
 * share the same name (common with template-cloned shapes).
 */
export const findShapeByName = (slide: SlideData, name: string): SlideShapeData | null => {
  for (const shape of slide[SLIDE_SHAPES]) {
    if (shape[SHAPE_SNAPSHOT].name === name) return shape;
  }
  return null;
};

/** Every shape on the slide whose `cNvPr@name` equals `name`. */
export const findShapesByName = (
  slide: SlideData,
  name: string,
): ReadonlyArray<SlideShapeData> =>
  slide[SLIDE_SHAPES].filter((s) => s[SHAPE_SNAPSHOT].name === name);

/** Every shape on the slide of the given kind. */
export const findShapesByKind = (
  slide: SlideData,
  kind: ShapeKind,
): ReadonlyArray<SlideShapeData> =>
  slide[SLIDE_SHAPES].filter((s) => s[SHAPE_SNAPSHOT].kind === kind);

/**
 * Walks every slide and returns the first shape whose name matches.
 * Useful for "find the logo placeholder anywhere in the deck."
 */
export const findShapeInPresentation = (
  pres: PresentationData,
  name: string,
): SlideShapeData | null => {
  for (const slide of getSlides(pres)) {
    const hit = findShapeByName(slide, name);
    if (hit !== null) return hit;
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

/**
 * Reads back the fill choice on the shape's `<p:spPr>`. Returns:
 *
 *   - `{ kind: 'solid', color: '#RRGGBB' }` for a solid sRGB fill.
 *   - `{ kind: 'solid', color: 'scheme:accent1' }` for a scheme color.
 *   - `{ kind: 'gradient' }` / `'pattern'` / `'image'` for those choices
 *     (without breaking out their parameters — call the dedicated
 *     setter to overwrite).
 *   - `{ kind: 'none' }` for `<a:noFill>`.
 *   - `{ kind: 'inherit' }` when no fill choice is present on this
 *     shape (it inherits from the layout / master placeholder).
 */
export type ShapeFill =
  | { readonly kind: 'solid'; readonly color: string }
  | { readonly kind: 'gradient' }
  | { readonly kind: 'pattern' }
  | { readonly kind: 'image' }
  | { readonly kind: 'none' }
  | { readonly kind: 'inherit' };

export const getShapeFill = (shape: SlideShapeData): ShapeFill => {
  const spPrName = qname('p', 'spPr', NS.pml);
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], spPrName);
  if (!spPr) return { kind: 'inherit' };
  for (const c of spPr.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    switch (c.name.localName) {
      case 'noFill':
        return { kind: 'none' };
      case 'solidFill': {
        // Look for the immediate color choice; report sRGB verbatim,
        // scheme colors as "scheme:<token>".
        for (const inner of c.children) {
          if (inner.kind !== 'element' || inner.name.namespaceURI !== NS.dml) continue;
          if (inner.name.localName === 'srgbClr') {
            const val = getAttrValue(inner, qname('', 'val', ''));
            if (val !== null) return { kind: 'solid', color: `#${val.toUpperCase()}` };
          }
          if (inner.name.localName === 'schemeClr') {
            const val = getAttrValue(inner, qname('', 'val', ''));
            if (val !== null) return { kind: 'solid', color: `scheme:${val}` };
          }
        }
        return { kind: 'solid', color: '' };
      }
      case 'gradFill':
        return { kind: 'gradient' };
      case 'pattFill':
        return { kind: 'pattern' };
      case 'blipFill':
        return { kind: 'image' };
    }
  }
  return { kind: 'inherit' };
};

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

// Rebuild shape handles entirely — used when the shape count changes
// (e.g. removeShape). Existing SlideShapeData identities are dropped;
// SHAPE_SLIDE back-pointers stay consistent because the SlideData
// reference is preserved.
const rebuildShapesFromDocument = (slide: SlideData): void => {
  const fresh = readSlidePart(slide[SLIDE_DOCUMENT].root);
  slide[SLIDE_PART] = fresh;
  const shapes: SlideShapeData[] = [];
  for (const snap of fresh.shapes) {
    shapes.push({
      [SHAPE_SLIDE]: slide,
      [SHAPE_ELEMENT]: snap.element,
      [SHAPE_SNAPSHOT]: snap,
    });
  }
  slide[SLIDE_SHAPES] = shapes;
};

// ---------------------------------------------------------------------------
// Shape mutation — geometry.

const NAME_TX_BODY_FN = qname('p', 'txBody', NS.pml);

const requireSpPr = (shape: SlideShapeData): XmlElement => {
  const kind = shape[SHAPE_SNAPSHOT].kind;
  if (kind !== 'shape' && kind !== 'picture' && kind !== 'connector') {
    throw new Error(`fill/stroke is not supported on ${kind} shapes`);
  }
  const spPrName = qname('p', 'spPr', NS.pml);
  const el = shape[SHAPE_ELEMENT];
  let spPr = firstChildElement(el, spPrName);
  if (spPr === null) {
    spPr = { kind: 'element', name: spPrName, attrs: [], prefixDecls: new Map(), children: [] };
    el.children.push(spPr);
  }
  return spPr;
};

const requireTxBody = (shape: SlideShapeData): XmlElement => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'shape') {
    throw new Error(
      `text operations require a shape kind; ${shape[SHAPE_SNAPSHOT].kind} is not text-bearing`,
    );
  }
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY_FN);
  if (txBody === null) {
    throw new Error(`shape "${shape[SHAPE_SNAPSHOT].name}" has no <p:txBody>`);
  }
  return txBody;
};

const commitAndRefresh = (shape: SlideShapeData): void => {
  commitSlideData(shape[SHAPE_SLIDE]);
  refreshSlideData(shape[SHAPE_SLIDE]);
};

/** Sets the shape's position in EMU. Companion to `setShapeSize`. */
export const setShapePosition = (shape: SlideShapeData, x: Emu, y: Emu): void => {
  writePosition(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind, x, y);
  commitAndRefresh(shape);
};

/** Sets the shape's size in EMU. */
export const setShapeSize = (shape: SlideShapeData, w: Emu, h: Emu): void => {
  writeSize(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind, w, h);
  commitAndRefresh(shape);
};

/**
 * Sets the shape's rotation in degrees (positive clockwise). Values are
 * normalized into `[0, 360)`; pass `0` to clear an existing rotation.
 */
export const setShapeRotation = (shape: SlideShapeData, degrees: number): void => {
  writeRotation(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind, degrees);
  commitAndRefresh(shape);
};

/** Sets the shape's flip flags. Properties default to current state when omitted. */
export const setShapeFlip = (
  shape: SlideShapeData,
  options: { horizontal?: boolean; vertical?: boolean },
): void => {
  writeFlip(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind, options);
  commitAndRefresh(shape);
};

// ---------------------------------------------------------------------------
// Shape mutation — fill / stroke.

/** Sets a solid fill on the shape (color in `#RRGGBB` or scheme token). */
export const setShapeFill = (shape: SlideShapeData, color: string): void => {
  setSolidFill(requireSpPr(shape), color);
  commitAndRefresh(shape);
};

/**
 * Sets a linear gradient fill on the shape. Stops must lie in `[0, 1]`;
 * `angleDeg` defaults to `90` (top → bottom).
 *
 * Example: red → blue top-to-bottom:
 *
 *   setShapeGradientFill(shape, {
 *     stops: [{ offset: 0, color: '#FF0000' }, { offset: 1, color: '#0000FF' }],
 *     angleDeg: 90,
 *   });
 */
export const setShapeGradientFill = (
  shape: SlideShapeData,
  options: GradientFillOptions,
): void => {
  setGradientFill(requireSpPr(shape), options);
  commitAndRefresh(shape);
};

/**
 * Sets a picture fill on the shape, embedding `bytes` as a new media
 * part and replacing any prior fill choice on the shape's `<p:spPr>`.
 *
 * The image stretches to fill the shape (`<a:stretch><a:fillRect/>`).
 * Format is detected from magic bytes; pass `options.format` to
 * override (useful for SVG or unusual extensions).
 *
 * Throws if the format can't be detected and isn't provided explicitly,
 * or if the shape kind doesn't carry a `<p:spPr>` (e.g. groups).
 */
export const setShapeImageFill = (
  shape: SlideShapeData,
  bytes: Uint8Array,
  options: { format?: ImageFormat } = {},
): void => {
  const format = options.format ?? detectImageFormat(bytes);
  if (format === null) {
    throw new Error(
      'setShapeImageFill: could not detect image format. Pass options.format explicitly.',
    );
  }
  const contentType = contentTypeForFormat(format);
  const extension = extensionForFormat(format);
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];

  // Allocate /ppt/media/imageN.<ext> (shared with addSlideImage's
  // numbering — both feed off the same /ppt/media space).
  let nextN = 1;
  const mediaPattern = /^\/ppt\/media\/image(\d+)\./;
  for (const p of pkg.parts) {
    const m = p.name.match(mediaPattern);
    if (m?.[1] !== undefined) {
      const n = Number.parseInt(m[1], 10);
      if (Number.isFinite(n) && n >= nextN) nextN = n + 1;
    }
  }
  const newMediaName = partName(`/ppt/media/image${nextN}.${extension}`);
  setOpcDefault(pkg, extension, contentType);
  pkg.addPart(newMediaName, contentType, bytes);

  // Slide → image rel.
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const newRId = nextRelId(rels.items.map((r) => r.id));
  rels.items.push({
    id: newRId,
    type: REL_TYPES.image,
    target: `../media/image${nextN}.${extension}`,
    targetMode: 'Internal',
  });
  pkg.setRels(slide[SLIDE_PART_NAME], rels);

  // Replace the shape's fill choice with <a:blipFill>.
  const spPr = requireSpPr(shape);
  const FILL_CHOICES = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);
  spPr.children = spPr.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && FILL_CHOICES.has(c.name.localName)),
  );
  const blipName = qname('a', 'blip', NS.dml);
  const stretchName = qname('a', 'stretch', NS.dml);
  const fillRectName = qname('a', 'fillRect', NS.dml);
  const blipFillName = qname('a', 'blipFill', NS.dml);
  const blip = elem(blipName, { attrs: [attr(qname('r', 'embed', NS.officeDocRels), newRId)] });
  const stretch = elem(stretchName, { children: [elem(fillRectName)] });
  const blipFill = elem(blipFillName, { children: [blip, stretch] });
  // <a:blipFill> takes the same slot as <a:solidFill>; insert at the
  // current insertion index. We use the same heuristic as setSolidFill —
  // before <a:ln> / effectLst / scene3d / extLst.
  let insertAt = spPr.children.length;
  for (let i = 0; i < spPr.children.length; i++) {
    const c = spPr.children[i];
    if (c?.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (
      c.name.localName === 'ln' ||
      c.name.localName === 'effectLst' ||
      c.name.localName === 'effectDag' ||
      c.name.localName === 'scene3d' ||
      c.name.localName === 'sp3d' ||
      c.name.localName === 'extLst'
    ) {
      insertAt = i;
      break;
    }
  }
  spPr.children.splice(insertAt, 0, blipFill);
  commitAndRefresh(shape);
};

/** Sets `<a:noFill>` on the shape, leaving it transparent. */
export const setShapeNoFill = (shape: SlideShapeData): void => {
  setNoFillImpl(requireSpPr(shape));
  commitAndRefresh(shape);
};

/**
 * Removes any fill choice from the shape; it then inherits its fill
 * from the layout / master placeholder it descends from.
 */
export const clearShapeFill = (shape: SlideShapeData): void => {
  clearFillImpl(requireSpPr(shape));
  commitAndRefresh(shape);
};

/** Sets a solid-color outline on the shape. */
export const setShapeStroke = (
  shape: SlideShapeData,
  options: { color?: string; widthEmu?: number },
): void => {
  setSolidStroke(requireSpPr(shape), options as StrokeOptions);
  commitAndRefresh(shape);
};

/** Sets an explicit "no outline" on the shape. */
export const setShapeNoStroke = (shape: SlideShapeData): void => {
  setNoStrokeImpl(requireSpPr(shape));
  commitAndRefresh(shape);
};

/** Removes any outline override; the shape then inherits stroke from layout. */
export const clearShapeStroke = (shape: SlideShapeData): void => {
  clearStrokeImpl(requireSpPr(shape));
  commitAndRefresh(shape);
};

// ---------------------------------------------------------------------------
// Shape mutation — text.

/**
 * Replaces the shape's visible text with `value`. Newlines start a new
 * paragraph. Existing run/paragraph properties are preserved so font,
 * color, size, alignment, and bullet style stay intact.
 */
export const setShapeText = (
  shape: SlideShapeData,
  value: string,
  options: { bullets?: BulletStyle } = {},
): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'shape') {
    throw new Error(
      `setShapeText only works on text-bearing shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY_FN);
  if (txBody === null) {
    throw new Error(`shape "${shape[SHAPE_SNAPSHOT].name}" has no <p:txBody>`);
  }
  setTextBody(txBody, value);
  if (options.bullets !== undefined) {
    applyBulletToAllParagraphs(txBody, options.bullets);
  }
  commitAndRefresh(shape);
};

/** Sets the bullet style on every paragraph in the shape's text body. */
export const setShapeBullets = (shape: SlideShapeData, style: BulletStyle): void => {
  applyBulletToAllParagraphs(requireTxBody(shape), style);
  commitAndRefresh(shape);
};

/** Sets the horizontal alignment of every paragraph in the shape's text. */
export const setShapeAlignment = (
  shape: SlideShapeData,
  align: ParagraphAlignment,
): void => {
  applyAlignmentToAllParagraphs(requireTxBody(shape), align);
  commitAndRefresh(shape);
};

/**
 * Applies `format` to every run in the shape's text. Run-property
 * attributes not addressed by `format` are preserved, so partial
 * updates compose.
 */
export const setShapeTextFormat = (shape: SlideShapeData, format: TextFormat): void => {
  applyFormatToAllRuns(requireTxBody(shape), format);
  commitAndRefresh(shape);
};

// ---------------------------------------------------------------------------
// Per-run text accessors.
//
// Lets callers reach into a shape's text body to read or format a
// specific paragraph or run. `applyFormatToAllRuns` covers the bulk-edit
// case; these helpers cover "make this one word red."

const NAME_A_P = qname('a', 'p', NS.dml);
const NAME_A_R = qname('a', 'r', NS.dml);
const NAME_A_RPR = qname('a', 'rPr', NS.dml);
const NAME_A_T = qname('a', 't', NS.dml);

const paragraphsOf = (txBody: XmlElement): XmlElement[] =>
  txBody.children.filter(
    (c): c is XmlElement =>
      c.kind === 'element' &&
      c.name.namespaceURI === NAME_A_P.namespaceURI &&
      c.name.localName === 'p',
  );

const runsOf = (paragraph: XmlElement): XmlElement[] =>
  paragraph.children.filter(
    (c): c is XmlElement =>
      c.kind === 'element' &&
      c.name.namespaceURI === NAME_A_R.namespaceURI &&
      c.name.localName === 'r',
  );

const requireParagraph = (shape: SlideShapeData, paragraphIndex: number): XmlElement => {
  const txBody = requireTxBody(shape);
  const paragraphs = paragraphsOf(txBody);
  const paragraph = paragraphs[paragraphIndex];
  if (!paragraph) {
    throw new RangeError(
      `paragraph index ${paragraphIndex} out of range (have ${paragraphs.length})`,
    );
  }
  return paragraph;
};

const requireRun = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): XmlElement => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const runs = runsOf(paragraph);
  const run = runs[runIndex];
  if (!run) {
    throw new RangeError(
      `run index ${runIndex} out of range in paragraph ${paragraphIndex} (have ${runs.length})`,
    );
  }
  return run;
};

const ensureRPr = (run: XmlElement): XmlElement => {
  const existing = firstChildElement(run, NAME_A_RPR);
  if (existing !== null) return existing;
  // `<a:rPr>` is the first child of `<a:r>` per the schema.
  const fresh = elem(NAME_A_RPR);
  run.children.unshift(fresh);
  return fresh;
};

const readRunText = (run: XmlElement): string => {
  const tEl = firstChildElement(run, NAME_A_T);
  if (tEl === null) return '';
  let out = '';
  for (const child of tEl.children) {
    if (child.kind === 'text' || child.kind === 'cdata') out += child.data;
  }
  return out;
};

const writeRunText = (run: XmlElement, value: string): void => {
  let tEl = firstChildElement(run, NAME_A_T);
  if (tEl === null) {
    tEl = elem(NAME_A_T);
    run.children.push(tEl);
  }
  tEl.children = [{ kind: 'text', data: value }];
};

/** Number of paragraphs in the shape's text body. Throws for non-text shapes. */
export const getShapeParagraphCount = (shape: SlideShapeData): number =>
  paragraphsOf(requireTxBody(shape)).length;

/**
 * Number of text runs in the given paragraph. Throws on out-of-range
 * paragraph index or non-text shapes.
 */
export const getShapeRunCount = (shape: SlideShapeData, paragraphIndex: number): number =>
  runsOf(requireParagraph(shape, paragraphIndex)).length;

/** Visible text of a single run. */
export const getShapeRunText = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): string => readRunText(requireRun(shape, paragraphIndex, runIndex));

/**
 * Sets the text of a single run. Existing rPr (font, size, color, ...)
 * is preserved — only the visible characters change.
 */
export const setShapeRunText = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
  text: string,
): void => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  writeRunText(run, text);
  commitAndRefresh(shape);
};

/**
 * Applies `format` to a single run. Run-property attributes not
 * addressed by `format` are preserved — partial updates compose.
 *
 * Example: bold the second word of the first paragraph:
 *
 *   setShapeRunFormat(shape, 0, 1, { bold: true, color: '#FF0000' });
 */
export const setShapeRunFormat = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
  format: TextFormat,
): void => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  const rPr = ensureRPr(run);
  applyRunFormatInternal(rPr, format);
  commitAndRefresh(shape);
};

/**
 * Sets an external hyperlink on every run in the shape's text. Allocates
 * (or reuses) a `hyperlink` relationship on the slide's `.rels`. Pass
 * `null` to clear.
 */
export const setShapeHyperlink = (shape: SlideShapeData, url: string | null): void => {
  const slide = shape[SHAPE_SLIDE];
  const txBody = requireTxBody(shape);
  if (url === null) {
    applyHyperlinkToAllRuns(txBody, null);
  } else {
    const pkg = slide[INTERNAL_PACKAGE];
    const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
    const existing = rels.items.find(
      (r) => r.type === REL_TYPES.hyperlink && r.target === url && r.targetMode === 'External',
    );
    const rId =
      existing?.id ??
      (() => {
        const nextId = nextRelId(rels.items.map((r) => r.id));
        rels.items.push({
          id: nextId,
          type: REL_TYPES.hyperlink,
          target: url,
          targetMode: 'External',
        });
        pkg.setRels(slide[SLIDE_PART_NAME], rels);
        return nextId;
      })();
    applyHyperlinkToAllRuns(txBody, rId);
  }
  commitAndRefresh(shape);
};

// ---------------------------------------------------------------------------
// Shape mutation — removal.

/**
 * Removes the shape from its slide's shape tree. Subsequent property
 * reads on this handle reflect the stale snapshot — discard it after.
 *
 * Removing a picture does NOT delete the underlying media part — it
 * may be referenced from other slides.
 */
export const removeShape = (shape: SlideShapeData): void => {
  const slide = shape[SHAPE_SLIDE];
  const doc = slide[SLIDE_DOCUMENT];
  const cSld = firstChildElement(doc.root, qname('p', 'cSld', NS.pml));
  if (!cSld) return;
  const spTree = firstChildElement(cSld, qname('p', 'spTree', NS.pml));
  if (!spTree) return;
  const idx = spTree.children.indexOf(shape[SHAPE_ELEMENT]);
  if (idx < 0) return;
  spTree.children.splice(idx, 1);
  commitSlideData(slide);
  rebuildShapesFromDocument(slide);
};

// ---------------------------------------------------------------------------
// Slide-level shape authoring.
//
// Each `addXxx` builds an XML element via an internal builder, appends
// it to the slide's `<p:spTree>`, commits, rebuilds the typed view, and
// returns the new SlideShapeData.

const requireSpTree = (slide: SlideData): XmlElement => {
  const cSld = firstChildElement(slide[SLIDE_DOCUMENT].root, NAME_CSLD);
  if (!cSld) throw new Error('slide has no <p:cSld>');
  const spTree = firstChildElement(cSld, NAME_SP_TREE);
  if (!spTree) throw new Error('slide has no <p:spTree>');
  return spTree;
};

const nextShapeId = (slide: SlideData): number => {
  let maxId = 0;
  for (const s of slide[SLIDE_PART].shapes) {
    if (s.id > maxId) maxId = s.id;
  }
  return Math.max(maxId, 1) + 1;
};

const appendAndReturnNewShape = (slide: SlideData, child: XmlElement): SlideShapeData => {
  const spTree = requireSpTree(slide);
  spTree.children.push(child);
  commitSlideData(slide);
  const previousLength = slide[SLIDE_SHAPES].length;
  rebuildShapesFromDocument(slide);
  const created = slide[SLIDE_SHAPES][previousLength];
  if (!created) throw new Error('appendShape: post-condition failed');
  return created;
};

/**
 * Adds a free-form text box to the slide. Returns the new shape.
 *
 * The box is a plain rectangle with no fill or outline carrying one
 * paragraph with one run. The shape id is allocated as one more than
 * the current max id.
 */
export const addSlideTextBox = (
  slide: SlideData,
  opts: { x: Emu; y: Emu; w: Emu; h: Emu; text: string; name?: string },
): SlideShapeData => {
  const sp = buildTextBox({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
    text: opts.text,
  });
  return appendAndReturnNewShape(slide, sp);
};

/**
 * Adds a preset shape (rectangle, ellipse, arrow, ...) to the slide.
 * Optional `text` seeds a single run.
 */
export const addSlideShape = (
  slide: SlideData,
  opts: {
    preset: PresetShape | string;
    x: Emu;
    y: Emu;
    w: Emu;
    h: Emu;
    text?: string;
    textAnchor?: 'l' | 'ctr' | 'r' | 't' | 'b';
    name?: string;
  },
): SlideShapeData => {
  const sp = buildShape({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    preset: opts.preset,
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
    ...(opts.text !== undefined ? { text: opts.text } : {}),
    ...(opts.textAnchor !== undefined ? { textAnchor: opts.textAnchor } : {}),
  });
  return appendAndReturnNewShape(slide, sp);
};

/** Adds a straight-line connector between two points. */
export const addSlideLine = (
  slide: SlideData,
  opts: {
    from: { x: Emu; y: Emu };
    to: { x: Emu; y: Emu };
    color?: string;
    widthEmu?: number;
    name?: string;
  },
): SlideShapeData => {
  const cxn = buildConnector({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    from: opts.from,
    to: opts.to,
    ...(opts.color !== undefined ? { color: opts.color } : {}),
    ...(opts.widthEmu !== undefined ? { widthEmu: opts.widthEmu } : {}),
  });
  return appendAndReturnNewShape(slide, cxn);
};

/**
 * Adds a table to the slide. Cells render as plain text with default
 * theme-aware styling; `firstRow` / `bandRow` flags drive PowerPoint's
 * banded-header look unless options say otherwise.
 */
export const addSlideTable = (
  slide: SlideData,
  opts: {
    x: Emu;
    y: Emu;
    w: Emu;
    h: Emu;
    rows: ReadonlyArray<ReadonlyArray<string>>;
    colWidths?: ReadonlyArray<Emu>;
    rowHeights?: ReadonlyArray<Emu>;
    firstRow?: boolean;
    bandRow?: boolean;
    name?: string;
  },
): SlideShapeData => {
  const frame = buildTable({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
    rows: opts.rows,
    ...(opts.colWidths !== undefined ? { colWidths: opts.colWidths } : {}),
    ...(opts.rowHeights !== undefined ? { rowHeights: opts.rowHeights } : {}),
    ...(opts.firstRow !== undefined ? { firstRow: opts.firstRow } : {}),
    ...(opts.bandRow !== undefined ? { bandRow: opts.bandRow } : {}),
  });
  return appendAndReturnNewShape(slide, frame);
};

/**
 * Adds a picture to the slide from raw bytes. Returns the new shape.
 *
 * Allocates a `/ppt/media/imageN.<ext>` part, registers a Content_Types
 * Default if the extension isn't yet covered, allocates a slide→image
 * rel, and appends a `<p:pic>` element to the slide's `<p:spTree>`.
 *
 * Format is detected from magic bytes; pass `opts.format` to override.
 */
export const addSlideImage = (
  slide: SlideData,
  bytes: Uint8Array,
  opts: { x: Emu; y: Emu; w: Emu; h: Emu; format?: ImageFormat; name?: string },
): SlideShapeData => {
  const pkg = slide[INTERNAL_PACKAGE];
  const format = opts.format ?? detectImageFormat(bytes);
  if (format === null) {
    throw new Error('addSlideImage: could not detect image format. Pass options.format explicitly.');
  }
  const contentType = contentTypeForFormat(format);
  const extension = extensionForFormat(format);

  let nextN = 1;
  const mediaPattern = /^\/ppt\/media\/image(\d+)\./;
  for (const p of pkg.parts) {
    const m = p.name.match(mediaPattern);
    if (m?.[1] !== undefined) {
      const n = Number.parseInt(m[1], 10);
      if (Number.isFinite(n) && n >= nextN) nextN = n + 1;
    }
  }
  const newMediaName = partName(`/ppt/media/image${nextN}.${extension}`);

  const hasDefault = pkg.contentTypes.defaults.some(
    (d) => d.extension.toLowerCase() === extension,
  );
  if (!hasDefault) {
    pkg.contentTypes.defaults.push({ extension, contentType });
  }
  pkg.addPart(newMediaName, contentType, bytes);

  const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const newRId = nextRelId(rels.items.map((r) => r.id));
  rels.items.push({
    id: newRId,
    type: REL_TYPES.image,
    target: `../media/image${nextN}.${extension}`,
    targetMode: 'Internal',
  });
  pkg.setRels(slide[SLIDE_PART_NAME], rels);

  const pic = buildPicture({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    rEmbed: newRId,
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
  });
  return appendAndReturnNewShape(slide, pic);
};

// ---------------------------------------------------------------------------
// Slide-level background + transition.

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
      c.kind === 'element' &&
      c.name.namespaceURI === NS.pml &&
      c.name.localName === 'transition',
  );
  if (!transition) return null;
  const speed = getAttrValue(transition, qname('', 'spd', '')) as
    | 'slow'
    | 'med'
    | 'fast'
    | null;
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

const setSlideBackgroundXml = (
  slide: SlideData,
  configure: (bgPr: XmlElement) => void,
): void => {
  const cSld = firstChildElement(slide[SLIDE_DOCUMENT].root, NAME_CSLD);
  if (!cSld) throw new Error('slide has no <p:cSld>');
  const bgName = qname('p', 'bg', NS.pml);
  const bgPrName = qname('p', 'bgPr', NS.pml);
  let bg = firstChildElement(cSld, bgName);
  if (bg === null) {
    bg = { kind: 'element', name: bgName, attrs: [], prefixDecls: new Map(), children: [] };
    cSld.children.unshift(bg);
  }
  bg.children = [];
  const bgPr: XmlElement = {
    kind: 'element',
    name: bgPrName,
    attrs: [],
    prefixDecls: new Map(),
    children: [],
  };
  bg.children.push(bgPr);
  configure(bgPr);
  commitSlideData(slide);
  refreshSlideData(slide);
};

/**
 * Reads back the slide's current background. Returns a discriminated
 * union mirroring `getShapeFill`'s shape, plus `inherit` when no
 * `<p:bg>` element is present (the slide picks up its background from
 * the layout / master).
 */
export type SlideBackground =
  | { readonly kind: 'solid'; readonly color: string }
  | { readonly kind: 'gradient' }
  | { readonly kind: 'pattern' }
  | { readonly kind: 'image' }
  | { readonly kind: 'inherit' };

export const getSlideBackground = (slide: SlideData): SlideBackground => {
  const cSld = firstChildElement(slide[SLIDE_DOCUMENT].root, NAME_CSLD);
  if (!cSld) return { kind: 'inherit' };
  const bg = firstChildElement(cSld, qname('p', 'bg', NS.pml));
  if (!bg) return { kind: 'inherit' };
  const bgPr = firstChildElement(bg, qname('p', 'bgPr', NS.pml));
  if (!bgPr) return { kind: 'inherit' };
  for (const c of bgPr.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    switch (c.name.localName) {
      case 'solidFill': {
        for (const inner of c.children) {
          if (inner.kind !== 'element' || inner.name.namespaceURI !== NS.dml) continue;
          if (inner.name.localName === 'srgbClr') {
            const val = getAttrValue(inner, qname('', 'val', ''));
            if (val !== null) return { kind: 'solid', color: `#${val.toUpperCase()}` };
          }
          if (inner.name.localName === 'schemeClr') {
            const val = getAttrValue(inner, qname('', 'val', ''));
            if (val !== null) return { kind: 'solid', color: `scheme:${val}` };
          }
        }
        return { kind: 'solid', color: '' };
      }
      case 'gradFill':
        return { kind: 'gradient' };
      case 'pattFill':
        return { kind: 'pattern' };
      case 'blipFill':
        return { kind: 'image' };
    }
  }
  return { kind: 'inherit' };
};

/** Sets a solid fill on the slide's background. */
export const setSlideBackground = (slide: SlideData, color: string): void => {
  setSlideBackgroundXml(slide, (bgPr) => setSolidFill(bgPr, color));
};

/** Clears any explicit slide background, restoring layout inheritance. */
export const clearSlideBackground = (slide: SlideData): void => {
  const cSld = firstChildElement(slide[SLIDE_DOCUMENT].root, NAME_CSLD);
  if (!cSld) return;
  cSld.children = cSld.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'bg'),
  );
  commitSlideData(slide);
  refreshSlideData(slide);
};

// ---------------------------------------------------------------------------
// Speaker notes.

const findNotesPartName = (slide: SlideData): PartName | null => {
  const rels = slide[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;
  const notesRel = rels.items.find((r) => r.type === REL_TYPES.notesSlide);
  if (!notesRel) return null;
  return notesRel.target.startsWith('/')
    ? partName(notesRel.target)
    : resolveTarget(slide[SLIDE_PART_NAME], notesRel.target);
};

/**
 * Returns the slide's speaker notes (`null` if none). Pulls plain text
 * from the `body` placeholder; multi-line notes use `\n`.
 */
export const getSlideNotes = (slide: SlideData): string | null => {
  const notesPartName = findNotesPartName(slide);
  if (notesPartName === null) return null;
  const part = slide[INTERNAL_PACKAGE].getPart(notesPartName);
  if (part === null) return null;
  const root = parseXml(decode(part.data)).root;
  const cSld = firstChildElement(root, NAME_CSLD);
  if (!cSld) return null;
  const spTree = firstChildElement(cSld, NAME_SP_TREE);
  if (!spTree) return null;
  for (const child of spTree.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.pml) continue;
    if (child.name.localName !== 'sp') continue;
    const nvSpPr = firstChildElement(child, qname('p', 'nvSpPr', NS.pml));
    if (!nvSpPr) continue;
    const nvPr = firstChildElement(nvSpPr, qname('p', 'nvPr', NS.pml));
    if (!nvPr) continue;
    const ph = firstChildElement(nvPr, qname('p', 'ph', NS.pml));
    if (!ph) continue;
    const txBody = firstChildElement(child, qname('p', 'txBody', NS.pml));
    if (!txBody) continue;
    const lines: string[] = [];
    for (const p of txBody.children) {
      if (p.kind !== 'element' || p.name.namespaceURI !== NS.dml || p.name.localName !== 'p') {
        continue;
      }
      let line = '';
      for (const r of p.children) {
        if (r.kind !== 'element' || r.name.namespaceURI !== NS.dml || r.name.localName !== 'r') {
          continue;
        }
        for (const tElement of r.children) {
          if (
            tElement.kind === 'element' &&
            tElement.name.namespaceURI === NS.dml &&
            tElement.name.localName === 't'
          ) {
            for (const tc of tElement.children) {
              if (tc.kind === 'text') line += tc.data;
            }
          }
        }
      }
      lines.push(line);
    }
    return lines.join('\n');
  }
  return null;
};

/**
 * Sets the slide's speaker notes. Creates the `notesSlide` part and
 * wires up the rels (slide ↔ notesSlide ↔ notesMaster) on first call;
 * subsequent calls just replace the body placeholder text.
 */
export const setSlideNotes = (slide: SlideData, value: string): void => {
  const pkg = slide[INTERNAL_PACKAGE];
  const notesPartName = findNotesPartName(slide);
  if (notesPartName !== null) {
    const part = pkg.getPart(notesPartName);
    if (part === null) throw new Error(`notes rel points at missing part ${notesPartName}`);
    const doc = parseXml(decode(part.data));
    const cSld = firstChildElement(doc.root, NAME_CSLD);
    if (!cSld) throw new Error('notesSlide has no <p:cSld>');
    const spTree = firstChildElement(cSld, NAME_SP_TREE);
    if (!spTree) throw new Error('notesSlide has no <p:spTree>');
    for (const child of spTree.children) {
      if (child.kind !== 'element' || child.name.namespaceURI !== NS.pml) continue;
      if (child.name.localName !== 'sp') continue;
      const nvSpPr = firstChildElement(child, qname('p', 'nvSpPr', NS.pml));
      if (!nvSpPr) continue;
      const nvPr = firstChildElement(nvSpPr, qname('p', 'nvPr', NS.pml));
      if (!nvPr) continue;
      const ph = firstChildElement(nvPr, qname('p', 'ph', NS.pml));
      if (!ph) continue;
      const txBody = firstChildElement(child, qname('p', 'txBody', NS.pml));
      if (!txBody) continue;
      setTextBody(txBody, value);
      part.data = encode(serializeXml(doc));
      return;
    }
    throw new Error('notesSlide has no body placeholder to fill');
  }

  // Create a new notesSlide part.
  const notesMasterPart = pkg.parts.find((p) => p.contentType.endsWith('notesMaster+xml'));
  let nextN = 1;
  const pattern = /^\/ppt\/notesSlides\/notesSlide(\d+)\.xml$/;
  for (const p of pkg.parts) {
    const m = p.name.match(pattern);
    if (m?.[1] !== undefined) {
      const n = Number.parseInt(m[1], 10);
      if (Number.isFinite(n) && n >= nextN) nextN = n + 1;
    }
  }
  const notesName = partName(`/ppt/notesSlides/notesSlide${nextN}.xml`);
  const doc = buildEmptyNotesSlide(value);
  pkg.addPart(
    notesName,
    'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
    encode(serializeXml(doc)),
  );

  const notesRels = emptyRels();
  const slideBase = slide[SLIDE_PART_NAME].split('/').pop() ?? 'slide.xml';
  notesRels.items.push({
    id: 'rId1',
    type: REL_TYPES.slide,
    target: `../slides/${slideBase}`,
    targetMode: 'Internal',
  });
  if (notesMasterPart) {
    const notesMasterBase = notesMasterPart.name.split('/').pop() ?? 'notesMaster1.xml';
    notesRels.items.push({
      id: 'rId2',
      type: REL_TYPES.notesMaster,
      target: `../notesMasters/${notesMasterBase}`,
      targetMode: 'Internal',
    });
  }
  pkg.setRels(notesName, notesRels);

  const slideRels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const existingIds = slideRels.items.map((r) => r.id);
  let n = 1;
  while (existingIds.includes(`rId${n}`)) n++;
  slideRels.items.push({
    id: `rId${n}`,
    type: REL_TYPES.notesSlide,
    target: `../notesSlides/notesSlide${nextN}.xml`,
    targetMode: 'Internal',
  });
  pkg.setRels(slide[SLIDE_PART_NAME], slideRels);
};

// ---------------------------------------------------------------------------
// Shape image replacement.

// ---------------------------------------------------------------------------
// Slide size.

/**
 * Width × height of the slide canvas, in EMU. `type` is PowerPoint's
 * aspect-ratio hint (`screen4x3`, `screen16x9`, ...); the actual size
 * is always `width` × `height`.
 */
export interface SlideSize {
  readonly width: Emu;
  readonly height: Emu;
  readonly type?: string;
}

/** Returns the slide canvas size, or `null` if `presentation.xml` omits it. */
export const getSlideSize = (pres: PresentationData): SlideSize | null => {
  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (presPart === null) return null;
  const root = parseXml(decode(presPart.data)).root;
  const model = readPresentationPart(root);
  if (model.slideSize === null) return null;
  return {
    width: model.slideSize.cx as Emu,
    height: model.slideSize.cy as Emu,
    ...(model.slideSize.type !== undefined ? { type: model.slideSize.type } : {}),
  };
};

const NAME_SLD_SZ_FN = qname('p', 'sldSz', NS.pml);
const ATTR_CX = qname('', 'cx', '');
const ATTR_CY = qname('', 'cy', '');
const ATTR_TYPE = qname('', 'type', '');
const NAME_SLD_ID_LST_FN = qname('p', 'sldIdLst', NS.pml);

/**
 * Sets the slide canvas size. Creates `<p:sldSz>` when absent, replaces
 * its attributes when present. The `type` hint is preserved as given.
 *
 * Schema ordering: `<p:sldSz>` follows `<p:sldIdLst>` per ECMA-376
 * §19.2.1.26; we insert at the correct position when bootstrapping.
 */
export const setSlideSize = (pres: PresentationData, opts: SlideSize): void => {
  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) throw new Error('presentation.xml is missing');
  const doc = parseXml(decode(presPart.data));

  let sldSz = firstChildElement(doc.root, NAME_SLD_SZ_FN);
  if (sldSz === null) {
    sldSz = elem(NAME_SLD_SZ_FN);
    const sldIdLst = firstChildElement(doc.root, NAME_SLD_ID_LST_FN);
    if (sldIdLst !== null) {
      const idx = doc.root.children.indexOf(sldIdLst);
      doc.root.children.splice(idx + 1, 0, sldSz);
    } else {
      doc.root.children.push(sldSz);
    }
  }

  sldSz.attrs = [attr(ATTR_CX, String(opts.width)), attr(ATTR_CY, String(opts.height))];
  if (opts.type !== undefined) sldSz.attrs.push(attr(ATTR_TYPE, opts.type));

  presPart.data = encode(serializeXml(doc));
};

import { emu as emuValue } from './units.ts';

/** 10in × 7.5in (`screen4x3`). */
export const SLIDE_SIZE_4_3: SlideSize = {
  width: emuValue(9144000),
  height: emuValue(6858000),
  type: 'screen4x3',
};

/** 13.333in × 7.5in (`screen16x9`) — Office 2013+ default. */
export const SLIDE_SIZE_16_9: SlideSize = {
  width: emuValue(12192000),
  height: emuValue(6858000),
  type: 'screen16x9',
};

/** 13.333in × 8.33in (`screen16x10`). */
export const SLIDE_SIZE_16_10: SlideSize = {
  width: emuValue(12192000),
  height: emuValue(7620000),
  type: 'screen16x10',
};

// ---------------------------------------------------------------------------
// Comments.
//
// Legacy schema (ECMA-376 Part 1 §19.4):
//   * One package-level `/ppt/commentAuthors.xml` holds every author.
//   * One `/ppt/comments/comment{N}.xml` per slide that has comments;
//     N matches the slide's part name.
//   * Slide rels reference the slide's comments part; presentation rels
//     reference the author list.
//
// Authors are deduped by (name, initials). `idx` allocation is per-author
// monotonic; we read each author's `lastIdx` and bump it on add.

const COMMENT_AUTHORS_PART_NAME = partName('/ppt/commentAuthors.xml');
const COMMENT_AUTHORS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml';
const COMMENTS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.comments+xml';

const slideNumberFromPartName = (name: PartName): number => {
  const m = name.match(/^\/ppt\/slides\/slide(\d+)\.xml$/);
  if (!m?.[1]) {
    throw new Error(`comments: cannot derive slide number from ${name}`);
  }
  return Number.parseInt(m[1], 10);
};

const commentsPartNameForSlide = (slide: SlideData): PartName => {
  const slideN = slideNumberFromPartName(slide[SLIDE_PART_NAME]);
  return partName(`/ppt/comments/comment${slideN}.xml`);
};

const loadAuthorList = (pkg: OpcPackage): CommentAuthor[] => {
  const part = pkg.getPart(COMMENT_AUTHORS_PART_NAME);
  if (part === null) return [];
  const list = readCommentAuthorList(parseXml(decode(part.data)).root);
  return list.authors.slice();
};

const writeAuthorList = (pkg: OpcPackage, authors: ReadonlyArray<CommentAuthor>): void => {
  const doc = buildCommentAuthorListDoc(authors);
  const bytes = encode(serializeXml(doc));
  const existing = pkg.getPart(COMMENT_AUTHORS_PART_NAME);
  if (existing !== null) {
    existing.data = bytes;
    return;
  }
  pkg.addPart(COMMENT_AUTHORS_PART_NAME, COMMENT_AUTHORS_CONTENT_TYPE, bytes);
  // presentation → commentAuthors rel.
  const presRels = pkg.getRels(PRES_PART_NAME) ?? emptyRels();
  const exists = presRels.items.some(
    (r) => r.type === REL_TYPES.commentAuthors && r.target.endsWith('commentAuthors.xml'),
  );
  if (!exists) {
    presRels.items.push({
      id: nextRelId(presRels.items.map((r) => r.id)),
      type: REL_TYPES.commentAuthors,
      target: 'commentAuthors.xml',
      targetMode: 'Internal',
    });
    pkg.setRels(PRES_PART_NAME, presRels);
  }
};

const loadCommentsForSlide = (slide: SlideData): SlideComment[] => {
  const pkg = slide[INTERNAL_PACKAGE];
  const partNameValue = commentsPartNameForSlide(slide);
  const part = pkg.getPart(partNameValue);
  if (part === null) return [];
  const list = readCommentList(parseXml(decode(part.data)).root);
  return list.comments.slice();
};

const writeCommentsForSlide = (
  slide: SlideData,
  comments: ReadonlyArray<SlideComment>,
): void => {
  const pkg = slide[INTERNAL_PACKAGE];
  const commentsName = commentsPartNameForSlide(slide);

  if (comments.length === 0) {
    // Drop the comments part + slide → comments rel when no comments
    // remain. Leaves an empty part orphaned otherwise.
    if (pkg.getPart(commentsName) !== null) {
      pkg.removePart(commentsName);
    }
    const slideRels = pkg.getRels(slide[SLIDE_PART_NAME]);
    if (slideRels !== null) {
      const before = slideRels.items.length;
      slideRels.items = slideRels.items.filter((r) => r.type !== REL_TYPES.comments);
      if (slideRels.items.length !== before) {
        pkg.setRels(slide[SLIDE_PART_NAME], slideRels);
      }
    }
    return;
  }

  const doc = buildCommentListDoc(comments);
  const bytes = encode(serializeXml(doc));
  const existing = pkg.getPart(commentsName);
  if (existing !== null) {
    existing.data = bytes;
    return;
  }
  pkg.addPart(commentsName, COMMENTS_CONTENT_TYPE, bytes);

  const slideRels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const hasRel = slideRels.items.some((r) => r.type === REL_TYPES.comments);
  if (!hasRel) {
    const slideN = slideNumberFromPartName(slide[SLIDE_PART_NAME]);
    slideRels.items.push({
      id: nextRelId(slideRels.items.map((r) => r.id)),
      type: REL_TYPES.comments,
      target: `../comments/comment${slideN}.xml`,
      targetMode: 'Internal',
    });
    pkg.setRels(slide[SLIDE_PART_NAME], slideRels);
  }
};

const asCommentData = (slide: SlideData, snap: SlideComment, author: CommentAuthor): SlideCommentData => ({
  [COMMENT_SLIDE]: slide,
  [COMMENT_SNAPSHOT]: snap,
  author,
});

/**
 * Every author known to the package's `commentAuthors.xml`.
 * Returns an empty array when no author list exists.
 */
export const getCommentAuthors = (pres: PresentationData): ReadonlyArray<CommentAuthor> =>
  loadAuthorList(pres[INTERNAL_PACKAGE]);

/**
 * Returns every comment attached to the slide, with the author already
 * resolved. The list is read-only — use `addSlideComment` /
 * `removeSlideComment` to mutate.
 */
export const getSlideComments = (slide: SlideData): ReadonlyArray<SlideCommentData> => {
  const pkg = slide[INTERNAL_PACKAGE];
  const authors = loadAuthorList(pkg);
  const authorById = new Map<number, CommentAuthor>();
  for (const a of authors) authorById.set(a.id, a);

  const comments = loadCommentsForSlide(slide);
  const out: SlideCommentData[] = [];
  for (const snap of comments) {
    const author = authorById.get(snap.authorId);
    if (!author) {
      // Comment references an unknown author — surface a synthetic
      // placeholder rather than dropping the comment silently.
      out.push(
        asCommentData(slide, snap, {
          id: snap.authorId,
          name: '',
          initials: '',
          lastIdx: snap.idx,
          clrIdx: null,
        }),
      );
      continue;
    }
    out.push(asCommentData(slide, snap, author));
  }
  return out;
};

/**
 * Adds a comment to the slide. Returns the new comment handle.
 *
 * Author handling: if an author with the given `name`+`initials` already
 * exists in `commentAuthors.xml`, the existing record is reused (and its
 * `lastIdx` is bumped). Otherwise a new author is allocated. `initials`
 * defaults to the first character of `name`.
 *
 * `position` is in EMU; pass `null` to omit the `<p:pos>` element.
 * `date` defaults to the current time.
 */
export const addSlideComment = (
  slide: SlideData,
  opts: {
    author: { name: string; initials?: string };
    text: string;
    position?: CommentPosition | null;
    date?: Date;
  },
): SlideCommentData => {
  const pkg = slide[INTERNAL_PACKAGE];
  const initials =
    opts.author.initials ?? (opts.author.name.length > 0 ? opts.author.name.charAt(0) : '?');

  const authors = loadAuthorList(pkg);
  let author = authors.find((a) => a.name === opts.author.name && a.initials === initials);
  if (!author) {
    let maxId = -1;
    for (const a of authors) if (a.id > maxId) maxId = a.id;
    author = {
      id: maxId + 1,
      name: opts.author.name,
      initials,
      lastIdx: 0,
      clrIdx: null,
    };
    authors.push(author);
  }
  const newIdx = author.lastIdx + 1;
  // Bump lastIdx on the author for the persisted list.
  const updatedAuthor: CommentAuthor = { ...author, lastIdx: newIdx };
  const persistedAuthors = authors.map((a) => (a.id === author!.id ? updatedAuthor : a));
  writeAuthorList(pkg, persistedAuthors);

  const dt = (opts.date ?? new Date()).toISOString();
  const snap: SlideComment = {
    authorId: updatedAuthor.id,
    idx: newIdx,
    dt,
    text: opts.text,
    position: opts.position ?? null,
  };

  const comments = loadCommentsForSlide(slide);
  comments.push(snap);
  writeCommentsForSlide(slide, comments);

  return asCommentData(slide, snap, updatedAuthor);
};

/**
 * Removes the comment from its slide's comments part. If the comment
 * was the last one on the slide, the comments part and the
 * slide → comments rel are also removed. The author entry in
 * `commentAuthors.xml` is left intact (an author may have comments on
 * other slides).
 */
export const removeSlideComment = (comment: SlideCommentData): void => {
  const slide = comment[COMMENT_SLIDE];
  const target = comment[COMMENT_SNAPSHOT];
  const remaining = loadCommentsForSlide(slide).filter(
    (c) => !(c.authorId === target.authorId && c.idx === target.idx),
  );
  writeCommentsForSlide(slide, remaining);
};

// Accessors over CommentAuthor / SlideCommentData for tree-shake convenience.

export const getCommentAuthor = (comment: SlideCommentData): CommentAuthor => comment.author;
export const getCommentText = (comment: SlideCommentData): string =>
  comment[COMMENT_SNAPSHOT].text;
export const getCommentDate = (comment: SlideCommentData): string | null =>
  comment[COMMENT_SNAPSHOT].dt;
export const getCommentPosition = (comment: SlideCommentData): CommentPosition | null =>
  comment[COMMENT_SNAPSHOT].position;

// ---------------------------------------------------------------------------

/**
 * Replaces a picture's media with `bytes`. Same-format replacements
 * write in place; cross-format replacements allocate a new media part
 * and repoint the rel. The original geometry — crop, sizing, transform —
 * is preserved.
 */
export const setShapeImage = (
  shape: SlideShapeData,
  bytes: Uint8Array,
  options: { format?: ImageFormat } = {},
): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') {
    throw new Error(`setShapeImage only works on picture shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`);
  }
  const format = options.format ?? detectImageFormat(bytes);
  if (format === null) {
    throw new Error('setShapeImage: could not detect image format. Pass options.format explicitly.');
  }
  const rEmbed = getPictureEmbedRId(shape[SHAPE_ELEMENT]);
  if (rEmbed === null) {
    throw new Error(`picture "${shape[SHAPE_SNAPSHOT].name}" has no r:embed`);
  }
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (rels === null) throw new Error(`slide ${slide[SLIDE_PART_NAME]} has no rels`);
  const rel = rels.items.find((r) => r.id === rEmbed);
  if (!rel) throw new Error(`slide rels missing entry for r:embed="${rEmbed}"`);

  const mediaName = rel.target.startsWith('/')
    ? partName(rel.target)
    : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
  const newExtension = extensionForFormat(format);
  const newContentType = contentTypeForFormat(format);
  const dotIdx = mediaName.lastIndexOf('.');
  const currentExtension = dotIdx >= 0 ? mediaName.slice(dotIdx + 1).toLowerCase() : '';

  if (currentExtension === newExtension) {
    const part = pkg.getPart(mediaName);
    if (!part) throw new Error(`media part missing: ${mediaName}`);
    part.data = bytes;
    part.contentType = newContentType;
    return;
  }

  let nextN = 1;
  const mediaPathRegex = /^\/ppt\/media\/image(\d+)\./;
  for (const p of pkg.parts) {
    const m = p.name.match(mediaPathRegex);
    if (m?.[1] !== undefined) {
      const num = Number.parseInt(m[1], 10);
      if (Number.isFinite(num) && num >= nextN) nextN = num + 1;
    }
  }
  const newPartName = partName(`/ppt/media/image${nextN}.${newExtension}`);
  const hasDefault = pkg.contentTypes.defaults.some(
    (d) => d.extension.toLowerCase() === newExtension,
  );
  if (!hasDefault) {
    pkg.contentTypes.defaults.push({ extension: newExtension, contentType: newContentType });
  }
  pkg.addPart(newPartName, newContentType, bytes);
  rel.target = `../media/image${nextN}.${newExtension}`;
  pkg.setRels(slide[SLIDE_PART_NAME], rels);
};

// ---------------------------------------------------------------------------
// Shape click action — `<a:hlinkClick>` on the shape's cNvPr.
//
// Two flavors today: open a URL (External rel) or jump to another slide
// in this deck (Internal rel + `action="ppaction://hlinksldjump"`).
//
// PowerPoint also supports preset actions like `nextslide`, `prevslide`,
// `firstslide`, `lastslide`, but they're niche enough to defer until a
// concrete user need shows up.

/** What clicking the shape should do. */
export type ShapeClickAction =
  | { readonly kind: 'url'; readonly url: string }
  | { readonly kind: 'slide'; readonly slide: SlideData }
  | { readonly kind: 'nextSlide' }
  | { readonly kind: 'prevSlide' }
  | { readonly kind: 'firstSlide' }
  | { readonly kind: 'lastSlide' };

const NAME_HLINK_CLICK_FN = qname('a', 'hlinkClick', NS.dml);

// cNvPr lives at different paths depending on shape kind. Returns null
// for kinds we don't know how to navigate yet (groups, etc.).
const findCNvPr = (shape: SlideShapeData): XmlElement | null => {
  const root = shape[SHAPE_ELEMENT];
  const kind = shape[SHAPE_SNAPSHOT].kind;
  const wrapperName =
    kind === 'shape'
      ? 'nvSpPr'
      : kind === 'picture'
        ? 'nvPicPr'
        : kind === 'connector'
          ? 'nvCxnSpPr'
          : kind === 'graphicFrame'
            ? 'nvGraphicFramePr'
            : null;
  if (wrapperName === null) return null;
  const wrapper = firstChildElement(root, qname('p', wrapperName, NS.pml));
  if (!wrapper) return null;
  return firstChildElement(wrapper, qname('p', 'cNvPr', NS.pml));
};

const removeExistingHlinkClick = (cNvPr: XmlElement): void => {
  cNvPr.children = cNvPr.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'hlinkClick'),
  );
};

const findExistingHyperlinkRel = (
  rels: ReturnType<OpcPackage['getRels']>,
  url: string,
): string | null => {
  if (rels === null) return null;
  const existing = rels.items.find(
    (rl) => rl.type === REL_TYPES.hyperlink && rl.target === url && rl.targetMode === 'External',
  );
  return existing?.id ?? null;
};

/**
 * Sets (or clears) the click action on the shape. Side effects:
 *
 *   - For `kind: 'url'`, a `hyperlink` rel is added (or reused) on the
 *     slide's rels with `targetMode="External"`. `<a:hlinkClick r:id=…/>`
 *     points at it.
 *   - For `kind: 'slide'`, a `slide` rel is added pointing at the
 *     target slide's part. The `<a:hlinkClick>` carries
 *     `action="ppaction://hlinksldjump"`.
 *   - For the preset navigations (`nextSlide`, `prevSlide`, ...), no rel
 *     is allocated; just the `action` attribute carries the preset.
 *   - `null` removes any existing `<a:hlinkClick>`.
 *
 * The shape must be one of `shape | picture | connector | graphicFrame`.
 * Groups don't carry their own click action in our model.
 */
export const setShapeClickAction = (
  shape: SlideShapeData,
  action: ShapeClickAction | null,
): void => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) {
    throw new Error(
      `setShapeClickAction: ${shape[SHAPE_SNAPSHOT].kind} shape has no cNvPr to attach to`,
    );
  }

  removeExistingHlinkClick(cNvPr);

  if (action === null) {
    commitAndRefresh(shape);
    return;
  }

  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];

  let rId: string | null = null;
  let actionAttr: string | null = null;

  switch (action.kind) {
    case 'url': {
      const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
      const reused = findExistingHyperlinkRel(rels, action.url);
      if (reused !== null) {
        rId = reused;
      } else {
        const newId = nextRelId(rels.items.map((r) => r.id));
        rels.items.push({
          id: newId,
          type: REL_TYPES.hyperlink,
          target: action.url,
          targetMode: 'External',
        });
        pkg.setRels(slide[SLIDE_PART_NAME], rels);
        rId = newId;
      }
      break;
    }
    case 'slide': {
      const target = action.slide[SLIDE_PART_NAME];
      const targetBase = basename(target);
      const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
      const existing = rels.items.find(
        (rl) =>
          rl.type === REL_TYPES.slide &&
          rl.target === `../slides/${targetBase}` &&
          rl.targetMode === 'Internal',
      );
      if (existing) {
        rId = existing.id;
      } else {
        const newId = nextRelId(rels.items.map((r) => r.id));
        rels.items.push({
          id: newId,
          type: REL_TYPES.slide,
          target: `../slides/${targetBase}`,
          targetMode: 'Internal',
        });
        pkg.setRels(slide[SLIDE_PART_NAME], rels);
        rId = newId;
      }
      actionAttr = 'ppaction://hlinksldjump';
      break;
    }
    case 'nextSlide':
      actionAttr = 'ppaction://hlinkshowjump?jump=nextslide';
      break;
    case 'prevSlide':
      actionAttr = 'ppaction://hlinkshowjump?jump=previousslide';
      break;
    case 'firstSlide':
      actionAttr = 'ppaction://hlinkshowjump?jump=firstslide';
      break;
    case 'lastSlide':
      actionAttr = 'ppaction://hlinkshowjump?jump=lastslide';
      break;
  }

  const attrs = [] as Array<ReturnType<typeof attr>>;
  if (rId !== null) attrs.push(attr(qname('r', 'id', NS.officeDocRels), rId));
  else attrs.push(attr(qname('r', 'id', NS.officeDocRels), ''));
  if (actionAttr !== null) attrs.push(attr(qname('', 'action', ''), actionAttr));

  cNvPr.children.push(
    elem(NAME_HLINK_CLICK_FN, {
      attrs,
    }),
  );

  commitAndRefresh(shape);
};

// ---------------------------------------------------------------------------
// Charts.
//
// Authoring path for ChartML (`/ppt/charts/chart{N}.xml`) + the embedded
// `/ppt/embeddings/Microsoft_Excel_Worksheet{N}.xlsx` workbook that
// PowerPoint requires for the "Edit data" action to work. See plan §P9
// and §Risks for the scope constraints.
//
// Public surface is intentionally narrow: one `addSlideChart` entry point
// that takes a typed `ChartSpec`. The internal layer handles the chart
// XML, the embedded xlsx ZIP, and all the relationship wiring.

const CHART_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const EMBEDDED_XLSX_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const allocateChartIndex = (pkg: OpcPackage): number => {
  let next = 1;
  const re = /^\/ppt\/charts\/chart(\d+)\.xml$/;
  for (const p of pkg.parts) {
    const m = p.name.match(re);
    if (m?.[1] !== undefined) {
      const n = Number.parseInt(m[1], 10);
      if (Number.isFinite(n) && n >= next) next = n + 1;
    }
  }
  return next;
};

const NAME_GRAPHIC_FRAME = qname('p', 'graphicFrame', NS.pml);
const NAME_NV_GRAPHIC_FRAME_PR = qname('p', 'nvGraphicFramePr', NS.pml);
const NAME_C_NV_PR_FN = qname('p', 'cNvPr', NS.pml);
const NAME_C_NV_GRAPHIC_FRAME_PR = qname('p', 'cNvGraphicFramePr', NS.pml);
const NAME_NV_PR = qname('p', 'nvPr', NS.pml);
const NAME_XFRM = qname('p', 'xfrm', NS.pml);
const NAME_OFF = qname('a', 'off', NS.dml);
const NAME_EXT = qname('a', 'ext', NS.dml);
const NAME_GRAPHIC = qname('a', 'graphic', NS.dml);
const NAME_GRAPHIC_DATA = qname('a', 'graphicData', NS.dml);
const NAME_C_CHART = qname('c', 'chart', NS.chart);

const buildChartGraphicFrame = (opts: {
  id: number;
  name: string;
  x: Emu;
  y: Emu;
  w: Emu;
  h: Emu;
  rEmbed: string;
}): XmlElement => {
  const cNvPr = elem(NAME_C_NV_PR_FN, {
    attrs: [attr(qname('', 'id', ''), String(opts.id)), attr(qname('', 'name', ''), opts.name)],
  });
  const nvGraphicFramePr = elem(NAME_NV_GRAPHIC_FRAME_PR, {
    children: [cNvPr, elem(NAME_C_NV_GRAPHIC_FRAME_PR), elem(NAME_NV_PR)],
  });
  const off = elem(NAME_OFF, {
    attrs: [attr(qname('', 'x', ''), String(opts.x)), attr(qname('', 'y', ''), String(opts.y))],
  });
  const ext = elem(NAME_EXT, {
    attrs: [attr(qname('', 'cx', ''), String(opts.w)), attr(qname('', 'cy', ''), String(opts.h))],
  });
  const xfrm = elem(NAME_XFRM, { children: [off, ext] });
  const chartRef = elem(NAME_C_CHART, {
    prefixDecls: new Map([
      ['c', NS.chart],
      ['r', NS.officeDocRels],
    ]),
    attrs: [attr(qname('r', 'id', NS.officeDocRels), opts.rEmbed)],
  });
  const graphicData = elem(NAME_GRAPHIC_DATA, {
    attrs: [attr(qname('', 'uri', ''), NS.chart)],
    children: [chartRef],
  });
  const graphic = elem(NAME_GRAPHIC, { children: [graphicData] });
  return elem(NAME_GRAPHIC_FRAME, { children: [nvGraphicFramePr, xfrm, graphic] });
};

const setOpcDefault = (pkg: OpcPackage, extension: string, contentType: string): void => {
  const has = pkg.contentTypes.defaults.some((d) => d.extension.toLowerCase() === extension);
  if (!has) pkg.contentTypes.defaults.push({ extension, contentType });
};

/**
 * Adds a chart to the slide. Returns the new shape handle (kind
 * `graphicFrame`). Supported chart kinds today: `bar`, `column`,
 * `line`, `pie` — see `ChartSpec.kind`.
 *
 * Side effects:
 *
 *   - Allocates `/ppt/charts/chart{N}.xml` for the chart definition.
 *   - Allocates `/ppt/embeddings/Microsoft_Excel_Worksheet{N}.xlsx` as
 *     a placeholder workbook (single sheet, header row + one row per
 *     category). PowerPoint reads the inline `<c:strCache>` /
 *     `<c:numCache>` so the workbook is for "Edit data" only.
 *   - Slide → chart and chart → workbook rels are wired with fresh rIds.
 *   - `<a:graphicFrame>` is appended to the slide's `<p:spTree>`.
 *
 * Constraints:
 *
 *   - `pie` charts require exactly one series.
 *   - All series should have at most `categories.length` values; missing
 *     values are treated as blanks (gaps in the visualization).
 */
export const addSlideChart = (
  slide: SlideData,
  opts: {
    spec: ChartSpec;
    x: Emu;
    y: Emu;
    w: Emu;
    h: Emu;
    name?: string;
  },
): SlideShapeData => {
  const pkg = slide[INTERNAL_PACKAGE];
  const chartN = allocateChartIndex(pkg);
  const chartPartName = partName(`/ppt/charts/chart${chartN}.xml`);
  const xlsxPartName = partName(`/ppt/embeddings/Microsoft_Excel_Worksheet${chartN}.xlsx`);

  // Build the embedded xlsx bytes. Each row in the sheet corresponds to
  // one category; header row carries the series names.
  const xlsxRows = opts.spec.categories.map((label, i) => ({
    label,
    values: opts.spec.series.map((s) => s.values[i] ?? null),
  }));
  const xlsxBytes = buildEmbeddedXlsx(
    opts.spec.series.map((s) => s.name),
    xlsxRows,
  );

  // Build the chart XML and serialize.
  const chartDoc = buildChartSpaceDoc(opts.spec);
  const chartBytes = encode(serializeXml(chartDoc));

  // Add the chart part + its rel → embedded xlsx.
  pkg.addPart(chartPartName, CHART_CONTENT_TYPE, chartBytes);

  // The xlsx is a binary part; xlsx is already an OPC zip so we add a
  // Content_Types override (no Default, since `.xlsx` shouldn't override
  // unrelated archive entries even though there's only one such part
  // here in practice).
  pkg.addPart(xlsxPartName, EMBEDDED_XLSX_CONTENT_TYPE, xlsxBytes);

  // Make sure `.rels` is a recognized Default (it always is by the time
  // we get here, but be defensive for new packages).
  setOpcDefault(pkg, 'rels', 'application/vnd.openxmlformats-package.relationships+xml');

  const chartRels = emptyRels();
  chartRels.items.push({
    id: 'rId1',
    type: REL_TYPES.package,
    target: `../embeddings/Microsoft_Excel_Worksheet${chartN}.xlsx`,
    targetMode: 'Internal',
  });
  pkg.setRels(chartPartName, chartRels);

  // Slide → chart rel.
  const slideRels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const slideChartRId = nextRelId(slideRels.items.map((r) => r.id));
  slideRels.items.push({
    id: slideChartRId,
    type: REL_TYPES.chart,
    target: `../charts/chart${chartN}.xml`,
    targetMode: 'Internal',
  });
  pkg.setRels(slide[SLIDE_PART_NAME], slideRels);

  // Build and append the <p:graphicFrame> wrapper.
  const frame = buildChartGraphicFrame({
    id: nextShapeId(slide),
    name: opts.name ?? `Chart ${chartN}`,
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
    rEmbed: slideChartRId,
  });
  return appendAndReturnNewShape(slide, frame);
};

// Re-export chart types for consumers.
export type { ChartKind, ChartSeries, ChartSpec };

void textNode;

// ---------------------------------------------------------------------------
// Validator.

export type { IssueSeverity, ValidationIssue };

/**
 * Runs a set of lightweight invariant checks on the package and
 * returns the list of issues found. An empty array means the deck
 * passes every check.
 *
 * Catches the common authoring mistakes — missing presentation.xml,
 * dangling slide rels, slides without a layout, etc. — without
 * depending on a heavyweight XSD engine, so it runs identically in
 * Node and the browser.
 *
 * Use it as a pre-save sanity check, especially after orchestrating
 * lots of mutations against the same package. Higher-fidelity XSD
 * validation lives in the test harness (Layer 1) and stays Node-only.
 */
export const validatePresentation = (pres: PresentationData): ReadonlyArray<ValidationIssue> =>
  validatePresentationPackage(pres[INTERNAL_PACKAGE]);

// ---------------------------------------------------------------------------
// Picture cropping — `<a:srcRect>` inside the picture's `<p:blipFill>`.
//
// Percentages are 0-1 fractions per side, converted to ECMA-376's
// `ST_Percentage` units (1/1000 of a percent, so 0.25 → "25000"). Pass
// `null` to remove an existing crop.

/** Crop a picture by fraction of each side. Omitted sides default to 0. */
export interface ImageCrop {
  readonly left?: number;
  readonly top?: number;
  readonly right?: number;
  readonly bottom?: number;
}

const NAME_BLIP_FILL_FN = qname('p', 'blipFill', NS.pml);
const NAME_SRC_RECT_FN = qname('a', 'srcRect', NS.dml);
const NAME_BLIP_FN = qname('a', 'blip', NS.dml);
const ATTR_CROP_L = qname('', 'l', '');
const ATTR_CROP_T = qname('', 't', '');
const ATTR_CROP_R = qname('', 'r', '');
const ATTR_CROP_B = qname('', 'b', '');

const fractionToST = (n: number | undefined): string | null => {
  if (n === undefined || n === 0) return null;
  if (!Number.isFinite(n) || n < 0 || n >= 1) {
    throw new RangeError(`crop fraction must be in [0, 1), got ${n}`);
  }
  return String(Math.round(n * 100000));
};

/**
 * Sets (or clears) a `<a:srcRect>` on a picture shape, cropping the
 * embedded image by the given fraction on each side. Pass `null` to
 * remove an existing crop.
 *
 * Fractions are in `[0, 1)` per side. `{ left: 0.25 }` clips 25% off
 * the left edge; the visible image stretches to fill the original
 * frame. The shape's geometry (`<a:xfrm>`) is unchanged.
 */
export const setShapeImageCrop = (shape: SlideShapeData, crop: ImageCrop | null): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') {
    throw new Error(
      `setShapeImageCrop only works on picture shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const pic = shape[SHAPE_ELEMENT];
  const blipFill = firstChildElement(pic, NAME_BLIP_FILL_FN);
  if (!blipFill) throw new Error('picture has no <p:blipFill>');

  // Remove any existing srcRect first.
  blipFill.children = blipFill.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'srcRect'),
  );

  if (crop === null) {
    commitAndRefresh(shape);
    return;
  }

  const attrs: Array<ReturnType<typeof attr>> = [];
  const l = fractionToST(crop.left);
  const t = fractionToST(crop.top);
  const r = fractionToST(crop.right);
  const b = fractionToST(crop.bottom);
  if (l !== null) attrs.push(attr(ATTR_CROP_L, l));
  if (t !== null) attrs.push(attr(ATTR_CROP_T, t));
  if (r !== null) attrs.push(attr(ATTR_CROP_R, r));
  if (b !== null) attrs.push(attr(ATTR_CROP_B, b));

  // <a:srcRect> sits between <a:blip> and <a:stretch> per the schema.
  const srcRect = elem(NAME_SRC_RECT_FN, { attrs });
  const blipIdx = blipFill.children.findIndex(
    (c) =>
      c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'blip',
  );
  if (blipIdx === -1) {
    // No <a:blip>? Just prepend the srcRect.
    blipFill.children.unshift(srcRect);
  } else {
    blipFill.children.splice(blipIdx + 1, 0, srcRect);
  }
  commitAndRefresh(shape);
};

void NAME_BLIP_FN;

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
export const setShapeAnimation = (
  shape: SlideShapeData,
  opts: AnimationOptions,
): void => {
  const slide = shape[SHAPE_SLIDE];
  removeExistingTiming(slide);
  const spid = shape[SHAPE_SNAPSHOT].id;
  const timing = buildSingleEffectTiming(spid, opts);
  insertTimingAtEnd(slide, timing);
  commitSlideData(slide);
  refreshSlideData(slide);
};

/** Removes the slide's `<p:timing>` element entirely. */
export const clearSlideAnimations = (slide: SlideData): void => {
  removeExistingTiming(slide);
  commitSlideData(slide);
  refreshSlideData(slide);
};

void NAME_TIMING_FN;

// ---------------------------------------------------------------------------
// Slide title convenience.
//
// Most decks bind their title placeholder to `type="title"` or `type="ctrTitle"`
// (the latter is the centered hero title on a "Title Slide" layout).
// These two helpers cover ~90% of the "set the slide title" use case.

/**
 * Returns the slide's title text, or `null` if neither a `title` nor
 * a `ctrTitle` placeholder is present.
 */
export const getSlideTitle = (slide: SlideData): string | null => {
  const titleShape =
    findSlidePlaceholder(slide, 'title') ?? findSlidePlaceholder(slide, 'ctrTitle');
  if (titleShape === null) return null;
  return titleShape[SHAPE_SNAPSHOT].text ?? null;
};

/**
 * Sets the slide's title text. Looks for a `title` placeholder first,
 * falling back to `ctrTitle`. Throws if neither exists — the slide's
 * layout has no title slot.
 */
export const setSlideTitle = (slide: SlideData, title: string): void => {
  const titleShape =
    findSlidePlaceholder(slide, 'title') ?? findSlidePlaceholder(slide, 'ctrTitle');
  if (titleShape === null) {
    throw new Error('setSlideTitle: slide has no title / ctrTitle placeholder');
  }
  setShapeText(titleShape, title);
};
