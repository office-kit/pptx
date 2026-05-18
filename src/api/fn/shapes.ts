// Shape read & write operations: queries, geometry, fill/stroke,
// effects, text, paragraphs, runs, color resolution, removal, z-order,
// and slide-level shape authoring.

import {
  type ArrowOptions,
  type BulletStyle,
  type GlowOptions,
  type GradientFillOptions,
  type LineDash,
  type ParagraphAlignment,
  type PatternFillOptions,
  type Position,
  type ShadowOptions,
  type Size,
  type StrokeOptions,
  type TextFormat,
  applyAlignmentToAllParagraphs,
  applyBulletToAllParagraphs,
  applyBulletToParagraph,
  applyFormatToAllRuns,
  applyHyperlinkToAllRuns,
  applyRunFormat as applyRunFormatInternal,
  clearEffects as clearEffectsImpl,
  clearFill as clearFillImpl,
  clearStroke as clearStrokeImpl,
  replaceTokensInTree,
  readFlip,
  readPosition,
  readRotation,
  readSize,
  setFlip as writeFlip,
  setGlow,
  setGradientFill,
  setPatternFill,
  setShadow,
  setNoFill as setNoFillImpl,
  setNoStroke as setNoStrokeImpl,
  setPosition as writePosition,
  setRotation as writeRotation,
  setSize as writeSize,
  setSolidFill,
  setSolidStroke,
  setStrokeArrow,
  setStrokeDash,
  setTextBody,
} from '../../internal/drawingml/index.ts';
import type { Emu } from '../units.ts';
import {
  basename,
  contentTypeForFormat,
  detectImageFormat,
  emptyRels,
  extensionForFormat,
  type ImageFormat,
  nextRelId,
  partName,
  resolveTarget,
} from '../../internal/opc/index.ts';
import {
  REL_TYPES,
  type PresetShape,
  type ShapeKind,
  buildConnector,
  buildPicture,
  buildShape,
  buildTable,
  buildTextBox,
  readGroupChildren,
  readPresentationPart,
  readShapeTreeFromCsldRoot,
  readSlideLayoutPart,
} from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
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
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  SLIDE_DOCUMENT,
  SLIDE_PART_NAME,
  SLIDE_SHAPES,
  type SlideData,
  type SlideLayoutData,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import {
  PRES_PART_NAME,
  appendAndReturnNewShape,
  commitAndRefresh,
  commitSlideData,
  decode,
  nextShapeId,
  rebuildShapesFromDocument,
  refreshSlideData,
  requireSpPr,
  requireSpTree,
  requireTxBody,
  setOpcDefault,
} from './_helpers.ts';
import { getPresentationFonts, getPresentationTheme, type PresentationTheme } from './package.ts';
import { getSlides } from './slides.ts';
import { findCNvPr, hasShapeText, NAME_HLINK_CLICK_FN, type ShapeClickAction } from './embedded.ts';

const NAME_TX_BODY = qname('p', 'txBody', NS.pml);

// ---------------------------------------------------------------------------
// Slide-level reads.

/**
 * Shapes on a slide, in document order with group children flattened.
 */
export const getSlideShapes = (slide: SlideData): ReadonlyArray<SlideShapeData> =>
  slide[SLIDE_SHAPES];

/**
 * Rebinds the slide to a different layout. The slide's own content
 * (shapes, text, geometry) is preserved verbatim; only the
 * `slideLayout` rel is updated so PowerPoint re-renders with the new
 * layout's placeholder positions and theme.
 *
 * The new layout must already be a part of the package — pass one
 * returned by `getSlideLayouts(pres)` or `findSlideLayout(pres, name)`.
 */
export const setSlideLayout = (slide: SlideData, layout: SlideLayoutData): void => {
  const pkg = slide[INTERNAL_PACKAGE];
  const layoutPartName = layout[LAYOUT_PART_NAME];
  if (pkg.getPart(layoutPartName) === null) {
    throw new Error(`setSlideLayout: layout ${layoutPartName} not in package`);
  }
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const layoutBase = basename(layoutPartName);
  const newTarget = `../slideLayouts/${layoutBase}`;

  // Replace any existing slideLayout rel. Keep the same rId where
  // possible so other parts that already reference it stay valid.
  const existing = rels.items.find((r) => r.type === REL_TYPES.slideLayout);
  if (existing) {
    existing.target = newTarget;
  } else {
    rels.items.push({
      id: nextRelId(rels.items.map((r) => r.id)),
      type: REL_TYPES.slideLayout,
      target: newTarget,
      targetMode: 'Internal',
    });
  }
  pkg.setRels(slide[SLIDE_PART_NAME], rels);
};

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
export const findSlidePlaceholder = (slide: SlideData, type: string): SlideShapeData | null => {
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
 * Returns every placeholder shape on the slide whose text body is
 * empty. Useful for "spot the slots that still need filling" UIs
 * before a slide is published, and for validation hooks that warn
 * about empty slots.
 */
export const findEmptyPlaceholders = (slide: SlideData): ReadonlyArray<SlideShapeData> => {
  const out: SlideShapeData[] = [];
  for (const shape of slide[SLIDE_SHAPES]) {
    if (!isShapePlaceholder(shape)) continue;
    if (hasShapeText(shape)) continue;
    out.push(shape);
  }
  return out;
};

/**
 * Returns the union bounding box of a group of shapes, or `null`
 * when none of them have bounds. Useful for "select all and move
 * together" patterns where the caller needs a single rectangle
 * across the group.
 */
export const getShapesBounds = (shapes: ReadonlyArray<SlideShapeData>): ShapeBounds | null => {
  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;
  let found = false;
  for (const shape of shapes) {
    const b = getShapeBounds(shape);
    if (!b) continue;
    found = true;
    if (b.x < minX) minX = b.x;
    if (b.y < minY) minY = b.y;
    if (b.x + b.w > maxX) maxX = b.x + b.w;
    if (b.y + b.h > maxY) maxY = b.y + b.h;
  }
  if (!found) return null;
  return {
    x: minX as Emu,
    y: minY as Emu,
    w: (maxX - minX) as Emu,
    h: (maxY - minY) as Emu,
  };
};

/**
 * Translates every shape in `shapes` by `(dxEmu, dyEmu)`. Useful
 * for "move this group of shapes 1cm right" patterns without
 * looping yourself. Shapes without bounds are skipped silently.
 */
export const translateShapes = (
  shapes: ReadonlyArray<SlideShapeData>,
  dxEmu: number,
  dyEmu: number,
): void => {
  for (const shape of shapes) {
    const bounds = getShapeBounds(shape);
    if (bounds === null) continue;
    setShapePosition(shape, (bounds.x + dxEmu) as Emu, (bounds.y + dyEmu) as Emu);
  }
};

/**
 * Returns every slide in the deck that has at least one empty
 * placeholder shape. Built on `findEmptyPlaceholders`. Useful for
 * "which slides still need editorial attention?" pre-publish
 * audits.
 */
export const getSlidesWithEmptyPlaceholders = (
  pres: PresentationData,
): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (findEmptyPlaceholders(slide).length > 0) out.push(slide);
  }
  return out;
};

/**
 * Returns the first placeholder shape whose `<p:ph idx="...">`
 * matches `idx`, or `null` when none does. Real templates often
 * disambiguate same-type placeholders (e.g. two body slots) by
 * `idx`, so this is what you reach for when type-only lookup is
 * ambiguous.
 */
export const findSlidePlaceholderByIdx = (slide: SlideData, idx: number): SlideShapeData | null => {
  for (const shape of slide[SLIDE_SHAPES]) {
    if (shape[SHAPE_SNAPSHOT].placeholderIdx === idx) return shape;
  }
  return null;
};

/**
 * Returns every placeholder shape with the given `type`. Useful for
 * "two-content" / "comparison" layouts where multiple body
 * placeholders share a type and the caller needs to fill them all.
 * Like `findSlidePlaceholder`, omitted `<p:ph type>` is treated as
 * `body` per ECMA-376 §19.7.10.
 */
export const findSlidePlaceholders = (
  slide: SlideData,
  type: string,
): ReadonlyArray<SlideShapeData> => {
  const out: SlideShapeData[] = [];
  for (const shape of slide[SLIDE_SHAPES]) {
    const snap = shape[SHAPE_SNAPSHOT];
    if (snap.placeholderType === type) {
      out.push(shape);
      continue;
    }
    if (type === 'body' && snap.placeholderType === null && snap.placeholderIdx !== null) {
      out.push(shape);
    }
  }
  return out;
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

/**
 * Returns the shape with the given OOXML internal id (`cNvPr@id`), or
 * `null` when no such shape exists. Shape ids are unique within a
 * slide; pair with `getShapeId` to round-trip references that arrive
 * from external XML (e.g. animations, hyperlinks).
 */
export const findShapeById = (slide: SlideData, id: number): SlideShapeData | null => {
  for (const shape of slide[SLIDE_SHAPES]) {
    if (shape[SHAPE_SNAPSHOT].id === id) return shape;
  }
  return null;
};

/** Every shape on the slide whose `cNvPr@name` equals `name`. */
export const findShapesByName = (slide: SlideData, name: string): ReadonlyArray<SlideShapeData> =>
  slide[SLIDE_SHAPES].filter((s) => s[SHAPE_SNAPSHOT].name === name);

/**
 * First shape on the slide whose visible text matches `needle`
 * (substring or `RegExp`), or `null` when none does. Convenience
 * over `getSlideShapes(slide).find(...)` when the caller is hunting
 * for a label in a template ("find the box that says 'Q1'").
 */
export const findShapeByText = (
  slide: SlideData,
  needle: string | RegExp,
): SlideShapeData | null => {
  for (const shape of slide[SLIDE_SHAPES]) {
    const text = shape[SHAPE_SNAPSHOT].text;
    if (typeof needle === 'string' ? text.includes(needle) : needle.test(text)) {
      return shape;
    }
  }
  return null;
};

/**
 * Every shape on the slide whose visible text matches `needle`. Use
 * when more than one shape can share the same text (common with
 * cloned bullet templates) — multi-match variant of
 * `findShapeByText`.
 */
export const findShapesByText = (
  slide: SlideData,
  needle: string | RegExp,
): ReadonlyArray<SlideShapeData> => {
  const out: SlideShapeData[] = [];
  for (const shape of slide[SLIDE_SHAPES]) {
    const text = shape[SHAPE_SNAPSHOT].text;
    if (typeof needle === 'string' ? text.includes(needle) : needle.test(text)) {
      out.push(shape);
    }
  }
  return out;
};

/** Every shape on the slide of the given kind. */
export const findShapesByKind = (
  slide: SlideData,
  kind: ShapeKind,
): ReadonlyArray<SlideShapeData> =>
  slide[SLIDE_SHAPES].filter((s) => s[SHAPE_SNAPSHOT].kind === kind);

/**
 * Returns every shape on the slide whose `<a:prstGeom prst="…"/>`
 * token matches `preset`. Useful for diagram introspection: find all
 * `'leftArrow'` shapes for a workflow swap, replace every `'cloud'`
 * with `'rect'`, etc.
 *
 * Shapes without a preset (custGeom / pictures / charts / tables /
 * connectors / groups) are filtered out.
 */
export const findShapesByPreset = (
  slide: SlideData,
  preset: string,
): ReadonlyArray<SlideShapeData> => slide[SLIDE_SHAPES].filter((s) => getShapePreset(s) === preset);

/**
 * Returns the slide that owns `shape`. Useful when callers receive a
 * shape from an unfiltered walk (`getAllShapes`, `findShapeInPresentation`,
 * search results) and need to know which slide it's on.
 */
export const getShapeSlide = (shape: SlideShapeData): SlideData => shape[SHAPE_SLIDE];

/**
 * Returns the shape's current XML element as a string. Diagnostic
 * sibling of `getSlideXmlString`; useful for snapshot tests, bug
 * reports, and before/after dumps during transformations.
 */
export const getShapeXmlString = (shape: SlideShapeData): string =>
  serializeXml({
    kind: 'document',
    decl: null,
    root: shape[SHAPE_ELEMENT],
    prolog: [],
    epilog: [],
  });

/**
 * Returns the 0-based document-order index of `shape` on its slide,
 * or `-1` when the shape is stale (e.g. after a `removeShape` that
 * rebuilt the slide's shape list).
 */
export const getShapeIndex = (shape: SlideShapeData): number => {
  const shapes = shape[SHAPE_SLIDE][SLIDE_SHAPES];
  return shapes.indexOf(shape);
};

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
export const replaceTokensInSlide = (slide: SlideData, tokens: Record<string, string>): number => {
  const n = replaceTokensInTree(slide[SLIDE_DOCUMENT].root, tokens);
  if (n > 0) {
    commitSlideData(slide);
    refreshSlideData(slide);
  }
  return n;
};

// ---------------------------------------------------------------------------
// SlideShape-level reads.

export const getShapeKind = (shape: SlideShapeData): ShapeKind => shape[SHAPE_SNAPSHOT].kind;

export const getShapeId = (shape: SlideShapeData): number => shape[SHAPE_SNAPSHOT].id;

/**
 * Returns the preset-geometry token (`'rect'`, `'ellipse'`, `'star5'`,
 * `'rightArrow'`, ...) for shapes whose body carries a
 * `<a:prstGeom prst="…"/>`. Returns `null` for:
 *
 *   - non-`'shape'` kinds (pictures, connectors, group shapes, tables,
 *     charts — they have their own geometry tags or no geometry),
 *   - shapes using custom geometry (`<a:custGeom>`),
 *   - shapes whose preset is missing (malformed but possible).
 *
 * Useful for renderers / inspectors that want to draw a faithful
 * approximation of each shape without dropping to the raw XML.
 */
export const getShapePreset = (shape: SlideShapeData): string | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'shape' && shape[SHAPE_SNAPSHOT].kind !== 'connector')
    return null;
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const prstGeom = firstChildElement(spPr, qname('a', 'prstGeom', NS.dml));
  if (!prstGeom) return null;
  for (const a of prstGeom.attrs) {
    if (a.name.localName === 'prst') return a.value;
  }
  return null;
};

/**
 * Reads the preset's adjust-handle values (`<a:prstGeom><a:avLst>
 * <a:gd name="adj" fmla="val 30000"/></a:avLst>`) as a map from guide
 * name → numeric value. Per ECMA-376 §20.1.9.4, guides are stored
 * with a formula prefix — `val 12345` is a literal number, and the
 * other prefixes (`pin`, `+-`, etc.) compute from other guides. We
 * only surface the `val` form because other formulas reference the
 * preset's built-in guides and don't make sense without them.
 *
 * Returns an empty record when no adjust values are authored (the
 * shape paints at its preset defaults).
 */
export const getShapeAdjustValues = (shape: SlideShapeData): Record<string, number> => {
  const out: Record<string, number> = {};
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return out;
  const prstGeom = firstChildElement(spPr, qname('a', 'prstGeom', NS.dml));
  if (!prstGeom) return out;
  const avLst = firstChildElement(prstGeom, qname('a', 'avLst', NS.dml));
  if (!avLst) return out;
  for (const gd of avLst.children) {
    if (gd.kind !== 'element' || gd.name.namespaceURI !== NS.dml || gd.name.localName !== 'gd')
      continue;
    const name = getAttrValue(gd, qname('', 'name', ''));
    const fmla = getAttrValue(gd, qname('', 'fmla', ''));
    if (!name || !fmla) continue;
    const match = /^val\s+(-?\d+(?:\.\d+)?)$/.exec(fmla);
    if (!match) continue;
    const n = Number.parseFloat(match[1]!);
    if (Number.isFinite(n)) out[name] = n;
  }
  return out;
};

/**
 * Returns the highest `cNvPr@id` used by any shape on the slide,
 * or `0` when the slide has no shapes with non-negative ids.
 *
 * Useful when hand-rolling a custom shape and you need an id known
 * not to collide. The official allocator inside `addSlideShape` /
 * `addSlideTextBox` etc. already does this — call those instead
 * when you don't need a custom id.
 */
export const getMaxShapeId = (slide: SlideData): number => {
  let max = 0;
  for (const shape of slide[SLIDE_SHAPES]) {
    const id = shape[SHAPE_SNAPSHOT].id;
    if (id > max) max = id;
  }
  return max;
};

/**
 * Deck-wide sibling of `getMaxShapeId`. Returns the highest
 * `cNvPr@id` across every shape on every slide, or `0` when the
 * deck has no shapes.
 *
 * Note: shape ids are scoped to a slide in OOXML — collisions
 * across slides are fine. This helper is for the rare cases where
 * a caller wants a single id known to be higher than anything in
 * the deck (e.g. to keep ids monotonically increasing).
 */
export const getMaxShapeIdInPresentation = (pres: PresentationData): number => {
  let max = 0;
  for (const slide of getSlides(pres)) {
    for (const shape of slide[SLIDE_SHAPES]) {
      const id = shape[SHAPE_SNAPSHOT].id;
      if (id > max) max = id;
    }
  }
  return max;
};

/**
 * Returns the number of slide masters declared in the
 * presentation's `<p:sldMasterIdLst>`. Most decks use exactly one
 * master; multi-master decks come from templates that combine
 * brand variants (e.g. a corporate master + a sponsor master).
 *
 * Returns `0` if `presentation.xml` is missing.
 */
export const getSlideMasterCount = (pres: PresentationData): number => {
  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (presPart === null) return 0;
  const root = parseXml(decode(presPart.data)).root;
  const model = readPresentationPart(root);
  return model.slideMasters.length;
};

/**
 * Returns the package part name of every slide master declared in
 * `presentation.xml`, resolved through the presentation's `.rels`.
 * Sibling of `getSlideMasterCount` for downstream tooling that
 * needs the master URIs (e.g. byte-level diff, custom validators).
 *
 * Returns an empty array when `presentation.xml` or its `.rels`
 * are missing.
 */
/**
 * Returns the part name of the slide master a slide inherits from
 * (`/ppt/slideMasters/slideMaster1.xml`), or `null` when the slide
 * has no layout or its layout has no master rel.
 *
 * Useful for multi-master decks where different slides live under
 * different brand templates and the caller needs to scope theme /
 * fontScheme / clrMap lookups to the correct master.
 */
export const getSlideMasterPartName = (slide: SlideData): string | null => {
  const layout = getSlideLayout(slide);
  if (!layout) return null;
  const pkg = slide[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const layoutRels = pkg.getRels(layoutPartName);
  if (!layoutRels) return null;
  const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return null;
  return resolveTarget(layoutPartName, masterRel.target);
};

export const getSlideMasterPartNames = (pres: PresentationData): ReadonlyArray<string> => {
  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (presPart === null) return [];
  const root = parseXml(decode(presPart.data)).root;
  const model = readPresentationPart(root);
  const rels = pkg.getRels(PRES_PART_NAME);
  if (rels === null) return [];
  const out: string[] = [];
  for (const m of model.slideMasters) {
    const rel = rels.items.find((r) => r.id === m.rId);
    if (rel === undefined) continue;
    const resolved = rel.target.startsWith('/')
      ? partName(rel.target)
      : resolveTarget(PRES_PART_NAME, rel.target);
    out.push(resolved);
  }
  return out;
};

export const getShapeName = (shape: SlideShapeData): string => shape[SHAPE_SNAPSHOT].name;

/**
 * Renames the shape's `cNvPr@name`. The display name is what
 * PowerPoint shows in the Selection Pane and what `findShapeByName`
 * matches on. Empty strings are allowed (matches PowerPoint behavior).
 */
/**
 * Reads the shape's alt-text description (`<p:cNvPr descr="...">`).
 * Accessibility tools (screen readers, contrast checkers) and
 * PowerPoint's "Alt Text" pane look at this field. Returns `null`
 * when no description is set.
 */
export const getShapeDescription = (shape: SlideShapeData): string | null => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) return null;
  return getAttrValue(cNvPr, qname('', 'descr', ''));
};

/**
 * Sets the shape's alt-text description (`<p:cNvPr descr="...">`).
 * Pass `null` to clear. Important for accessibility — image shapes
 * and decorative graphics should carry a descr that conveys the
 * visual meaning to screen readers.
 */
export const setShapeDescription = (shape: SlideShapeData, description: string | null): void => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) {
    throw new Error(`setShapeDescription: ${shape[SHAPE_SNAPSHOT].kind} shape has no cNvPr`);
  }
  cNvPr.attrs = cNvPr.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'descr'),
  );
  if (description !== null && description !== '') {
    cNvPr.attrs.push(attr(qname('', 'descr', ''), description));
  }
  commitAndRefresh(shape);
};

/**
 * Reads the shape's alt-text title (`<p:cNvPr title="...">`).
 * PowerPoint surfaces this alongside `descr` in its Alt Text pane
 * as a short heading. Returns `null` when no title is set.
 */
export const getShapeAltTitle = (shape: SlideShapeData): string | null => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) return null;
  return getAttrValue(cNvPr, qname('', 'title', ''));
};

/**
 * Sets the shape's alt-text title (`<p:cNvPr title="...">`). Pass
 * `null` to clear. Distinct from `renameShape`, which writes the
 * `name` attribute used in the selection pane.
 */
export const setShapeAltTitle = (shape: SlideShapeData, title: string | null): void => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) {
    throw new Error(`setShapeAltTitle: ${shape[SHAPE_SNAPSHOT].kind} shape has no cNvPr`);
  }
  cNvPr.attrs = cNvPr.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'title'),
  );
  if (title !== null && title !== '') {
    cNvPr.attrs.push(attr(qname('', 'title', ''), title));
  }
  commitAndRefresh(shape);
};

/**
 * `true` when the shape's `<p:cNvPr hidden="1">` is set. Hidden
 * shapes are skipped by PowerPoint's renderer but stay in the
 * shape tree — useful for variant slides that toggle which boxes
 * are visible.
 */
export const isShapeHidden = (shape: SlideShapeData): boolean => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) return false;
  const v = getAttrValue(cNvPr, qname('', 'hidden', ''));
  return v === '1' || v === 'true';
};

/**
 * Sets or clears `<p:cNvPr hidden="...">` on the shape. Hidden
 * shapes remain in the document but PowerPoint doesn't render them.
 */
export const setShapeHidden = (shape: SlideShapeData, hidden: boolean): void => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) {
    throw new Error(`setShapeHidden: ${shape[SHAPE_SNAPSHOT].kind} shape has no cNvPr`);
  }
  cNvPr.attrs = cNvPr.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'hidden'),
  );
  if (hidden) cNvPr.attrs.push(attr(qname('', 'hidden', ''), '1'));
  commitAndRefresh(shape);
};

export const renameShape = (shape: SlideShapeData, newName: string): void => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) {
    throw new Error(`renameShape: ${shape[SHAPE_SNAPSHOT].kind} shape has no cNvPr to rename`);
  }
  cNvPr.attrs = cNvPr.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'name'),
  );
  cNvPr.attrs.push(attr(qname('', 'name', ''), newName));
  commitAndRefresh(shape);
};

export const getShapePlaceholderType = (shape: SlideShapeData): string | null =>
  shape[SHAPE_SNAPSHOT].placeholderType;

export const getShapePlaceholderIdx = (shape: SlideShapeData): number | null =>
  shape[SHAPE_SNAPSHOT].placeholderIdx;

/**
 * `true` when the shape carries `<p:nvSpPr><p:nvPr><p:ph>` — i.e. it
 * inherits from a layout/master placeholder. False for decorative
 * geometry the slide author dropped onto the canvas. Decoupled from
 * the more specific `getShapePlaceholderType` / `getShapePlaceholderIdx`
 * (either can be null on a real placeholder; together they identify it).
 */
export const isShapePlaceholder = (shape: SlideShapeData): boolean => {
  const snap = shape[SHAPE_SNAPSHOT];
  return snap.placeholderType !== null || snap.placeholderIdx !== null;
};

export const getShapeText = (shape: SlideShapeData): string => shape[SHAPE_SNAPSHOT].text;

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
 * Enumerates the shapes nested inside a `<p:grpSp>` group, one level
 * deep (nested groups come through as `kind: 'group'` themselves —
 * call this again on each one to recurse).
 *
 * Returns an empty array for non-group shapes. Each child carries its
 * own bounds in the group's *internal* coordinate system; pair with
 * `getGroupTransform` to project them onto the slide.
 */
export const getGroupChildren = (shape: SlideShapeData): ReadonlyArray<SlideShapeData> => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'group') return [];
  const children = readGroupChildren(shape[SHAPE_ELEMENT]);
  return children.map((child) => ({
    [SHAPE_SLIDE]: shape[SHAPE_SLIDE],
    [SHAPE_ELEMENT]: child.element,
    [SHAPE_SNAPSHOT]: child,
  }));
};

/**
 * Returns the group's slide-relative bounds (`outer`) and the internal
 * coordinate system the children's `<a:xfrm>` values live in
 * (`inner`). Renderers project a child point `(cx, cy)` onto the slide
 * with:
 *
 *   slideX = outer.x + (cx - inner.x) * (outer.w / inner.w)
 *   slideY = outer.y + (cy - inner.y) * (outer.h / inner.h)
 *
 * Returns `null` for non-group shapes or for groups whose
 * `<p:grpSpPr>` omits an `<a:xfrm>`.
 */
export const getGroupTransform = (
  shape: SlideShapeData,
): {
  readonly outer: ShapeBounds;
  readonly inner: ShapeBounds;
} | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'group') return null;
  const grpSpPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'grpSpPr', NS.pml));
  if (!grpSpPr) return null;
  const xfrm = firstChildElement(grpSpPr, qname('a', 'xfrm', NS.dml));
  if (!xfrm) return null;
  const off = firstChildElement(xfrm, qname('a', 'off', NS.dml));
  const ext = firstChildElement(xfrm, qname('a', 'ext', NS.dml));
  const chOff = firstChildElement(xfrm, qname('a', 'chOff', NS.dml));
  const chExt = firstChildElement(xfrm, qname('a', 'chExt', NS.dml));
  if (!off || !ext) return null;
  const parseAttr = (el: XmlElement, name: string): number | null => {
    const raw = getAttrValue(el, qname('', name, ''));
    if (raw === null) return null;
    const n = Number.parseInt(raw, 10);
    return Number.isFinite(n) ? n : null;
  };
  const ox = parseAttr(off, 'x');
  const oy = parseAttr(off, 'y');
  const ow = parseAttr(ext, 'cx');
  const oh = parseAttr(ext, 'cy');
  if (ox === null || oy === null || ow === null || oh === null) return null;
  // Per ECMA-376, `<a:chOff>/<a:chExt>` default to the same values as
  // `<a:off>/<a:ext>` when omitted (i.e. no internal-to-outer
  // transform).
  const ix = chOff ? (parseAttr(chOff, 'x') ?? ox) : ox;
  const iy = chOff ? (parseAttr(chOff, 'y') ?? oy) : oy;
  const iw = chExt ? (parseAttr(chExt, 'cx') ?? ow) : ow;
  const ih = chExt ? (parseAttr(chExt, 'cy') ?? oh) : oh;
  return {
    outer: { x: ox as Emu, y: oy as Emu, w: ow as Emu, h: oh as Emu },
    inner: { x: ix as Emu, y: iy as Emu, w: iw as Emu, h: ih as Emu },
  };
};

/**
 * Combined bounds — position + size in one object. Returns `null` when
 * the shape inherits both position and size from its layout (so the
 * `<a:xfrm>` element is absent or incomplete).
 */
export interface ShapeBounds {
  readonly x: Emu;
  readonly y: Emu;
  readonly w: Emu;
  readonly h: Emu;
}

export const getShapeBounds = (shape: SlideShapeData): ShapeBounds | null => {
  const pos = readPosition(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind);
  const size = readSize(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind);
  if (pos === null || size === null) return null;
  return {
    x: pos.x as Emu,
    y: pos.y as Emu,
    w: size.w as Emu,
    h: size.h as Emu,
  };
};

/**
 * Same as `getShapeBounds` but walks the placeholder inheritance chain
 * when the shape has no `<a:xfrm>` of its own:
 *
 *   1. The shape's own bounds.
 *   2. The matching placeholder on the slide's layout (by `<p:ph idx>`,
 *      falling back to `<p:ph type>`).
 *   3. The matching placeholder on the layout's slide master.
 *
 * This is what renderers want when they need to draw a placeholder
 * that the deck author left unsized — real templates only override
 * geometry per placeholder when it differs from the master.
 *
 * Returns `null` when none of the three levels carries explicit bounds.
 */
export const getShapeBoundsResolved = (
  pres: PresentationData,
  shape: SlideShapeData,
): ShapeBounds | null => {
  const direct = getShapeBounds(shape);
  if (direct) return direct;

  const slide = shape[SHAPE_SLIDE];
  const layout = getSlideLayout(slide);
  if (!layout) return null;

  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);

  const findInShapes = (
    shapes: ReadonlyArray<{
      placeholderIdx: number | null;
      placeholderType: string | null;
      element: XmlElement;
      kind: ShapeKind;
    }>,
  ): ShapeBounds | null => {
    let match = phIdx !== null ? shapes.find((s) => s.placeholderIdx === phIdx) : undefined;
    if (!match && phType !== null) {
      match = shapes.find((s) => s.placeholderType === phType);
    }
    if (!match) return null;
    const pos = readPosition(match.element, match.kind);
    const size = readSize(match.element, match.kind);
    if (pos === null || size === null) return null;
    return {
      x: pos.x as Emu,
      y: pos.y as Emu,
      w: size.w as Emu,
      h: size.h as Emu,
    };
  };

  const layoutHit = findInShapes(layout[LAYOUT_PART].shapes);
  if (layoutHit) return layoutHit;

  // Walk one level up: layout → slideMaster rel.
  const pkg = pres[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const layoutRels = pkg.getRels(layoutPartName);
  if (!layoutRels) return null;
  const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return null;
  const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
  if (!masterPart) return null;

  const masterRoot = parseXml(decode(masterPart.data)).root;
  const { shapes: masterShapes } = readShapeTreeFromCsldRoot(masterRoot, 'sldMaster');
  return findInShapes(masterShapes);
};

/**
 * Returns the center point of the shape's bounds in EMU, or `null`
 * when the shape has no `<a:xfrm>`. Convenience for layout
 * pipelines that compute alignment / overlap from center points.
 */
export const getShapeCenter = (
  shape: SlideShapeData,
): { readonly x: Emu; readonly y: Emu } | null => {
  const bounds = getShapeBounds(shape);
  if (bounds === null) return null;
  return {
    x: (bounds.x + Math.round(bounds.w / 2)) as Emu,
    y: (bounds.y + Math.round(bounds.h / 2)) as Emu,
  };
};

/**
 * `true` when point `(x, y)` (in EMU) lies inside the shape's
 * axis-aligned bounds. Closed on the top-left edge, open on the
 * bottom-right (standard half-open rectangle). Returns `false`
 * when the shape has no bounds.
 *
 * Useful for hit-testing in custom interaction handlers.
 */
export const pointInShape = (shape: SlideShapeData, x: number, y: number): boolean => {
  const bounds = getShapeBounds(shape);
  if (bounds === null) return false;
  return x >= bounds.x && x < bounds.x + bounds.w && y >= bounds.y && y < bounds.y + bounds.h;
};

/**
 * Returns every shape on the slide whose bounds contain `(x, y)`
 * (in EMU). Built on `pointInShape`. The list is in document
 * order, so callers can index by z-stack from front (last) to
 * back (first) if they want one-hit semantics.
 */
export const findShapesAtPoint = (
  slide: SlideData,
  x: number,
  y: number,
): ReadonlyArray<SlideShapeData> => slide[SLIDE_SHAPES].filter((s) => pointInShape(s, x, y));

/**
 * Moves the shape so its center sits at the slide canvas center.
 * Reads the presentation's slide size, then sets the shape's
 * position to `(slideWidth/2 - shapeWidth/2, slideHeight/2 - shapeHeight/2)`.
 *
 * No-op when the shape has no bounds or the presentation has no
 * configured slide size.
 */
export const centerShapeOnSlide = (shape: SlideShapeData): void => {
  const bounds = getShapeBounds(shape);
  if (bounds === null) return;
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) return;
  const presRoot = parseXml(decode(presPart.data)).root;
  const model = readPresentationPart(presRoot);
  if (model.slideSize === null) return;
  const newX = Math.round(model.slideSize.cx / 2 - bounds.w / 2) as Emu;
  const newY = Math.round(model.slideSize.cy / 2 - bounds.h / 2) as Emu;
  setShapePosition(shape, newX, newY);
};

/**
 * `true` when two shapes' axis-aligned bounding boxes overlap.
 * Returns `false` when either shape has no bounds. Doesn't account
 * for rotation — uses the raw `<a:xfrm>` rectangle, not the
 * visual bounding box after rotation.
 *
 * Useful for collision detection in custom layout pipelines.
 */
export const shapesOverlap = (a: SlideShapeData, b: SlideShapeData): boolean => {
  const ba = getShapeBounds(a);
  const bb = getShapeBounds(b);
  if (ba === null || bb === null) return false;
  return ba.x < bb.x + bb.w && ba.x + ba.w > bb.x && ba.y < bb.y + bb.h && ba.y + ba.h > bb.y;
};

/**
 * Sets both position and size in one call. Equivalent to calling
 * `setShapePosition` followed by `setShapeSize`, but commits the slide
 * just once.
 */
export const setShapeBounds = (shape: SlideShapeData, bounds: ShapeBounds): void => {
  writePosition(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind, bounds.x, bounds.y);
  writeSize(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind, bounds.w, bounds.h);
  commitAndRefresh(shape);
};

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

/**
 * Reads back the shape's stroke (`<a:ln>`). Returns:
 *
 *   - `{ kind: 'solid', color, widthEmu? }` for a solid-color outline.
 *   - `{ kind: 'none' }` when an `<a:noFill>` sits inside `<a:ln>`.
 *   - `{ kind: 'inherit' }` when no `<a:ln>` is present.
 */
export type ShapeStroke =
  | { readonly kind: 'solid'; readonly color: string; readonly widthEmu?: number }
  | { readonly kind: 'none' }
  | { readonly kind: 'inherit' };

/**
 * Convenience over `getShapeStroke(shape)`: returns the solid-
 * stroke color (`#RRGGBB` / `scheme:<token>`) or `null` when the
 * stroke is inherited / removed.
 */
export const getShapeStrokeColor = (shape: SlideShapeData): string | null => {
  const stroke = getShapeStroke(shape);
  return stroke.kind === 'solid' ? stroke.color : null;
};

/**
 * Convenience over `getShapeStroke(shape)`: returns the stroke
 * width in EMU when the stroke is solid and an explicit width is
 * set, or `null` otherwise.
 */
export const getShapeStrokeWidth = (shape: SlideShapeData): number | null => {
  const stroke = getShapeStroke(shape);
  return stroke.kind === 'solid' && stroke.widthEmu !== undefined ? stroke.widthEmu : null;
};

/**
 * Returns the shape's stroke color resolved to a concrete `#RRGGBB`:
 * scheme tokens are mapped through the deck's color scheme and
 * `<a:lumMod>` / `<a:tint>` / `<a:shade>` / etc. transform children
 * are applied. Returns `null` when the stroke isn't a solid color
 * (inherits / `noFill`) or when the color can't be resolved.
 *
 * Companion to `getShapeStrokeColor`, which surfaces only the raw
 * `#RRGGBB` / `scheme:<token>` string — fine for round-tripping but
 * wrong for rendering, because PowerPoint paints the *transformed*
 * color, not the base one.
 */
export const getShapeStrokeColorResolved = (
  pres: PresentationData,
  shape: SlideShapeData,
): string | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return null;
  const solid = firstChildElement(ln, qname('a', 'solidFill', NS.dml));
  if (!solid) return null;
  for (const inner of solid.children) {
    if (inner.kind !== 'element' || inner.name.namespaceURI !== NS.dml) continue;
    return resolveDrawingColor(inner, getPresentationTheme(pres));
  }
  return null;
};

/**
 * Reads the stroke's line cap style — `'rnd'` (round), `'sq'` (square),
 * `'flat'`, or `null` when the attribute isn't set. Per ECMA-376
 * §20.1.2.3.10 (`ST_LineCap`).
 */
export const getShapeStrokeCap = (shape: SlideShapeData): 'rnd' | 'sq' | 'flat' | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return null;
  const v = getAttrValue(ln, qname('', 'cap', ''));
  if (v === 'rnd' || v === 'sq' || v === 'flat') return v;
  return null;
};

/**
 * Reads the stroke's line join style — `'round'` / `'bevel'` / `'miter'`,
 * or `null` when no explicit join element is present. Maps from the
 * three child-element variants `<a:round/>`, `<a:bevel/>`, `<a:miter/>`.
 */
export const getShapeStrokeJoin = (shape: SlideShapeData): 'round' | 'bevel' | 'miter' | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return null;
  for (const c of ln.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (c.name.localName === 'round') return 'round';
    if (c.name.localName === 'bevel') return 'bevel';
    if (c.name.localName === 'miter') return 'miter';
  }
  return null;
};

/**
 * Reads the stroke's compound-line style (`<a:ln cmpd="…">`) — single,
 * double, triple, or thick/thin / thin/thick parallel lines. ECMA-376
 * §20.1.2.3.11 (`ST_CompoundLine`).
 */
export const getShapeStrokeCompound = (
  shape: SlideShapeData,
): 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri' | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return null;
  const v = getAttrValue(ln, qname('', 'cmpd', ''));
  if (v === 'sng' || v === 'dbl' || v === 'thickThin' || v === 'thinThick' || v === 'tri') return v;
  return null;
};

/**
 * Same as `getShapeStroke` but walks the layout → master placeholder
 * cascade when the shape itself reports `'inherit'`. First non-inherit
 * stroke layer wins.
 */
export const getShapeStrokeEffective = (
  pres: PresentationData,
  shape: SlideShapeData,
): ShapeStroke => {
  const own = getShapeStroke(shape);
  if (own.kind !== 'inherit') return own;

  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);
  if (phIdx === null && phType === null) return own;

  const layout = getSlideLayout(shape[SHAPE_SLIDE]);
  if (!layout) return own;

  const readStrokeFromSpPr = (el: XmlElement): ShapeStroke | null => {
    const spPr = firstChildElement(el, qname('p', 'spPr', NS.pml));
    if (!spPr) return null;
    const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
    if (!ln) return null;
    const wRaw = getAttrValue(ln, qname('', 'w', ''));
    const widthEmu = wRaw !== null ? Number.parseInt(wRaw, 10) : undefined;
    for (const c of ln.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      if (c.name.localName === 'noFill') return { kind: 'none' };
      if (c.name.localName === 'solidFill') {
        for (const inner of c.children) {
          if (inner.kind !== 'element' || inner.name.namespaceURI !== NS.dml) continue;
          if (inner.name.localName === 'srgbClr') {
            const val = getAttrValue(inner, qname('', 'val', ''));
            if (val !== null) {
              return {
                kind: 'solid',
                color: `#${val.toUpperCase()}`,
                ...(widthEmu !== undefined ? { widthEmu } : {}),
              };
            }
          }
          if (inner.name.localName === 'schemeClr') {
            const val = getAttrValue(inner, qname('', 'val', ''));
            if (val !== null) {
              return {
                kind: 'solid',
                color: `scheme:${val}`,
                ...(widthEmu !== undefined ? { widthEmu } : {}),
              };
            }
          }
        }
      }
    }
    return null;
  };

  const findPh = (
    shapes: ReadonlyArray<{
      placeholderIdx: number | null;
      placeholderType: string | null;
      element: XmlElement;
    }>,
  ): XmlElement | null => {
    let match = phIdx !== null ? shapes.find((s) => s.placeholderIdx === phIdx) : undefined;
    if (!match && phType !== null) match = shapes.find((s) => s.placeholderType === phType);
    return match?.element ?? null;
  };

  const layoutPh = findPh(layout[LAYOUT_PART].shapes);
  if (layoutPh) {
    const s = readStrokeFromSpPr(layoutPh);
    if (s) return s;
  }
  const pkg = pres[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const layoutRels = pkg.getRels(layoutPartName);
  if (!layoutRels) return own;
  const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return own;
  const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
  if (!masterPart) return own;
  const masterRoot = parseXml(decode(masterPart.data)).root;
  const { shapes: masterShapes } = readShapeTreeFromCsldRoot(masterRoot, 'sldMaster');
  const masterPh = findPh(masterShapes);
  if (masterPh) {
    const s = readStrokeFromSpPr(masterPh);
    if (s) return s;
  }
  return own;
};

export const getShapeStroke = (shape: SlideShapeData): ShapeStroke => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return { kind: 'inherit' };
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return { kind: 'inherit' };

  const wRaw = getAttrValue(ln, qname('', 'w', ''));
  const widthEmu = wRaw !== null ? Number.parseInt(wRaw, 10) : undefined;

  for (const c of ln.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (c.name.localName === 'noFill') return { kind: 'none' };
    if (c.name.localName === 'solidFill') {
      for (const inner of c.children) {
        if (inner.kind !== 'element' || inner.name.namespaceURI !== NS.dml) continue;
        if (inner.name.localName === 'srgbClr') {
          const val = getAttrValue(inner, qname('', 'val', ''));
          if (val !== null) {
            return {
              kind: 'solid',
              color: `#${val.toUpperCase()}`,
              ...(widthEmu !== undefined ? { widthEmu } : {}),
            };
          }
        }
        if (inner.name.localName === 'schemeClr') {
          const val = getAttrValue(inner, qname('', 'val', ''));
          if (val !== null) {
            return {
              kind: 'solid',
              color: `scheme:${val}`,
              ...(widthEmu !== undefined ? { widthEmu } : {}),
            };
          }
        }
      }
      return {
        kind: 'solid',
        color: '',
        ...(widthEmu !== undefined ? { widthEmu } : {}),
      };
    }
  }
  return { kind: 'inherit' };
};

/**
 * Convenience over `getShapeFill(shape)`: returns the solid-fill
 * color string (`#RRGGBB` or `scheme:<token>`) when the shape has
 * one, or `null` otherwise. Use when the caller only cares about
 * the color and doesn't need to distinguish "inherit" / "no fill" /
 * "gradient" / "pattern" / "image" from each other.
 */
export const getShapeFillColor = (shape: SlideShapeData): string | null => {
  const fill = getShapeFill(shape);
  return fill.kind === 'solid' ? fill.color : null;
};

/**
 * Returns the shape's solid fill resolved to a concrete `#RRGGBB`:
 * scheme tokens are mapped through the deck's color scheme and
 * `<a:lumMod>` / `<a:tint>` / `<a:shade>` / etc. transform children
 * are applied. Returns `null` when the fill isn't solid (gradient,
 * pattern, image, none, inherit) or when the color can't be resolved.
 *
 * Companion to `getShapeFillColor`, which surfaces only the raw
 * `#RRGGBB` / `scheme:<token>` string. Renderers and exporters that
 * need the color PowerPoint actually paints should call this.
 */
export const getShapeFillColorResolved = (
  pres: PresentationData,
  shape: SlideShapeData,
): string | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const solid = firstChildElement(spPr, qname('a', 'solidFill', NS.dml));
  if (!solid) return null;
  for (const inner of solid.children) {
    if (inner.kind !== 'element' || inner.name.namespaceURI !== NS.dml) continue;
    return resolveDrawingColor(inner, getPresentationTheme(pres));
  }
  return null;
};

/**
 * Same as `getShapeFill` but walks the layout → master placeholder
 * cascade when the shape itself reports `'inherit'`. Returns the first
 * non-inherit fill found, or `{ kind: 'inherit' }` when neither layer
 * supplies one. Useful for renderers that want the actual fill the
 * placeholder will paint with.
 */
export const getShapeFillEffective = (pres: PresentationData, shape: SlideShapeData): ShapeFill => {
  const own = getShapeFill(shape);
  if (own.kind !== 'inherit') return own;

  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);
  if (phIdx === null && phType === null) return own;

  const layout = getSlideLayout(shape[SHAPE_SLIDE]);
  if (!layout) return own;

  const readFillFromSpPr = (el: XmlElement): ShapeFill | null => {
    const spPr = firstChildElement(el, qname('p', 'spPr', NS.pml));
    if (!spPr) return null;
    for (const c of spPr.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      switch (c.name.localName) {
        case 'noFill':
          return { kind: 'none' };
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
    return null;
  };

  const findPh = (
    shapes: ReadonlyArray<{
      placeholderIdx: number | null;
      placeholderType: string | null;
      element: XmlElement;
    }>,
  ): XmlElement | null => {
    let match = phIdx !== null ? shapes.find((s) => s.placeholderIdx === phIdx) : undefined;
    if (!match && phType !== null) match = shapes.find((s) => s.placeholderType === phType);
    return match?.element ?? null;
  };

  const layoutPh = findPh(layout[LAYOUT_PART].shapes);
  if (layoutPh) {
    const f = readFillFromSpPr(layoutPh);
    if (f) return f;
  }

  const pkg = pres[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const layoutRels = pkg.getRels(layoutPartName);
  if (!layoutRels) return own;
  const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return own;
  const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
  if (!masterPart) return own;
  const masterRoot = parseXml(decode(masterPart.data)).root;
  const { shapes: masterShapes } = readShapeTreeFromCsldRoot(masterRoot, 'sldMaster');
  const masterPh = findPh(masterShapes);
  if (masterPh) {
    const f = readFillFromSpPr(masterPh);
    if (f) return f;
  }
  return own;
};

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
// Detailed gradient-fill reader. Companion to `getShapeFill`, which
// only reports the discriminated `kind`. Returns the full stop list +
// angle when the shape carries a `<a:gradFill>` of its own, or
// `null` for solid / pattern / image / none / inherited fills.
//
// Useful for renderers (preview generators, PDF exporters) that need
// to reproduce the gradient instead of substituting a placeholder.

export const NAME_A_GRAD_FILL = qname('a', 'gradFill', NS.dml);
export const NAME_A_GS_LST = qname('a', 'gsLst', NS.dml);
export const NAME_A_LIN = qname('a', 'lin', NS.dml);

export const readColorFromContainer = (parent: XmlElement): string | null => {
  for (const c of parent.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (c.name.localName === 'srgbClr') {
      const val = getAttrValue(c, qname('', 'val', ''));
      if (val !== null) return `#${val.toUpperCase()}`;
    }
    if (c.name.localName === 'schemeClr') {
      const val = getAttrValue(c, qname('', 'val', ''));
      if (val !== null) return `scheme:${val}`;
    }
  }
  return null;
};

/**
 * Returns the full gradient definition (`stops` + `angleDeg`) when the
 * shape's `<p:spPr>` carries an `<a:gradFill>`. Returns `null` for any
 * other fill kind, including `inherit` — the function does not walk the
 * layout / master cascade.
 */
export const getShapeGradientFill = (shape: SlideShapeData): GradientFillOptions | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const gradFill = firstChildElement(spPr, NAME_A_GRAD_FILL);
  if (!gradFill) return null;
  const gsLst = firstChildElement(gradFill, NAME_A_GS_LST);
  if (!gsLst) return null;
  const stops: Array<{ offset: number; color: string }> = [];
  for (const c of gsLst.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml || c.name.localName !== 'gs') {
      continue;
    }
    const posRaw = getAttrValue(c, qname('', 'pos', ''));
    if (posRaw === null) continue;
    const pos = Number.parseInt(posRaw, 10);
    if (!Number.isFinite(pos)) continue;
    const color = readColorFromContainer(c);
    if (color === null) continue;
    stops.push({ offset: pos / 100_000, color });
  }
  if (stops.length === 0) return null;
  // ECMA-376 §20.1.8.33: gradFill has either <a:lin> (linear) or <a:path>
  // (radial / rectangular / shape-following) as a child to describe the
  // direction. We surface both so renderers can faithfully reproduce
  // non-linear gradients.
  let angleDeg = 0;
  const lin = firstChildElement(gradFill, NAME_A_LIN);
  if (lin) {
    const angRaw = getAttrValue(lin, qname('', 'ang', ''));
    if (angRaw !== null) {
      const ang = Number.parseInt(angRaw, 10);
      if (Number.isFinite(ang)) angleDeg = ang / 60_000;
    }
  }
  const pathEl = firstChildElement(gradFill, qname('a', 'path', NS.dml));
  if (pathEl) {
    const p = getAttrValue(pathEl, qname('', 'path', ''));
    const pathVal: 'circle' | 'rect' | 'shape' | null =
      p === 'circle' || p === 'rect' || p === 'shape' ? p : null;
    if (pathVal) {
      let focus: GradientFillOptions['focus'];
      const fillToRect = firstChildElement(pathEl, qname('a', 'fillToRect', NS.dml));
      if (fillToRect) {
        const pct = (name: string): number | undefined => {
          const v = getAttrValue(fillToRect, qname('', name, ''));
          if (v === null) return undefined;
          let n = Number.parseFloat(v);
          if (!Number.isFinite(n)) return undefined;
          if (Math.abs(n) > 1) n = n / 100000;
          return n;
        };
        const l = pct('l') ?? 0.5;
        const t = pct('t') ?? 0.5;
        const r = pct('r') ?? 0.5;
        const b = pct('b') ?? 0.5;
        focus = { left: l, top: t, right: r, bottom: b };
      }
      return { stops, angleDeg, path: pathVal, ...(focus ? { focus } : {}) };
    }
  }
  return { stops, angleDeg };
};

// ---------------------------------------------------------------------------
// Shape mutation — geometry.

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
export const setShapeGradientFill = (shape: SlideShapeData, options: GradientFillOptions): void => {
  setGradientFill(requireSpPr(shape), options);
  commitAndRefresh(shape);
};

/**
 * Sets a preset pattern fill on the shape (e.g. `pct50`, `dkUpDiag`).
 *
 * `foreground` is the pattern stroke color; `background` fills behind
 * the pattern. Both accept `#RRGGBB`, bare `RRGGBB`, or scheme tokens
 * (`accent1`, `bg1`, ...).
 */
export const setShapePatternFill = (shape: SlideShapeData, options: PatternFillOptions): void => {
  setPatternFill(requireSpPr(shape), options);
  commitAndRefresh(shape);
};

/**
 * Reads back the pattern fill on a shape: returns the preset token
 * plus the foreground / background colors resolved against the theme.
 * Returns `null` when the shape has no `<a:pattFill>`.
 *
 * The preset string is the literal `ST_PresetPatternVal` token from
 * §20.1.10.49 — e.g. `'pct50'`, `'dkUpDiag'`, `'cross'`, `'wave'`.
 * Renderers can map it onto an SVG `<pattern>` definition.
 */
export const getShapePatternFill = (
  pres: PresentationData,
  shape: SlideShapeData,
): { preset: string; foreground: string; background: string } | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const pattFill = firstChildElement(spPr, qname('a', 'pattFill', NS.dml));
  if (!pattFill) return null;
  const preset = getAttrValue(pattFill, qname('', 'prst', '')) ?? 'pct50';
  const theme = getPresentationTheme(pres);
  const colorFrom = (parentName: string, fallback: string): string => {
    const parent = firstChildElement(pattFill, qname('a', parentName, NS.dml));
    if (!parent) return fallback;
    for (const c of parent.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      const hex = resolveDrawingColor(c, theme);
      if (hex) return hex;
    }
    return fallback;
  };
  return {
    preset,
    foreground: colorFrom('fgClr', '#000000'),
    background: colorFrom('bgClr', '#FFFFFF'),
  };
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
  const FILL_CHOICES = new Set([
    'noFill',
    'solidFill',
    'gradFill',
    'blipFill',
    'pattFill',
    'grpFill',
  ]);
  spPr.children = spPr.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        FILL_CHOICES.has(c.name.localName)
      ),
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

/** Reads back the shape's stroke dash style, or `null` if none. */
export const getShapeStrokeDash = (shape: SlideShapeData): LineDash | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return null;
  const prstDash = firstChildElement(ln, qname('a', 'prstDash', NS.dml));
  if (!prstDash) return null;
  const v = getAttrValue(prstDash, qname('', 'val', ''));
  return (v as LineDash | null) ?? null;
};

/**
 * Reads back the shape's arrowhead on one end of `<a:ln>`, or `null`
 * when no `<a:headEnd>` / `<a:tailEnd>` is present.
 */
export const getShapeStrokeArrow = (
  shape: SlideShapeData,
  end: 'head' | 'tail',
): ArrowOptions | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return null;
  const arr = firstChildElement(ln, qname('a', end === 'head' ? 'headEnd' : 'tailEnd', NS.dml));
  if (!arr) return null;
  const type = getAttrValue(arr, qname('', 'type', ''));
  if (!type) return null;
  const width = getAttrValue(arr, qname('', 'w', ''));
  const length = getAttrValue(arr, qname('', 'len', ''));
  const result: {
    type: ArrowOptions['type'];
    width?: 'sm' | 'med' | 'lg';
    length?: 'sm' | 'med' | 'lg';
  } = {
    type: type as ArrowOptions['type'],
  };
  if (width === 'sm' || width === 'med' || width === 'lg') result.width = width;
  if (length === 'sm' || length === 'med' || length === 'lg') result.length = length;
  return result;
};

/**
 * Sets the dash pattern for the shape's outline (`<a:prstDash>`). One
 * of ECMA-376's `ST_PresetLineDashVal` tokens:
 *
 *   `'solid'` | `'dot'` | `'dash'` | `'lgDash'` | `'dashDot'` |
 *   `'lgDashDot'` | `'lgDashDotDot'` | `'sysDash'` | `'sysDot'` |
 *   `'sysDashDot'` | `'sysDashDotDot'`
 *
 * Creates `<a:ln>` if absent. Pairs naturally with `setShapeStroke`:
 * users typically set a color + width first, then the dash.
 */
export const setShapeStrokeDash = (shape: SlideShapeData, dash: LineDash): void => {
  setStrokeDash(requireSpPr(shape), dash);
  commitAndRefresh(shape);
};

/**
 * Sets an arrowhead on one end of the shape's outline.
 *
 *   - `end: 'head'` writes `<a:headEnd>` (the start of the line).
 *   - `end: 'tail'` writes `<a:tailEnd>` (the end).
 *
 * Useful primarily on connector shapes added via `addSlideLine`.
 * `type: 'none'` clears the arrowhead.
 */
export const setShapeStrokeArrow = (
  shape: SlideShapeData,
  end: 'head' | 'tail',
  options: ArrowOptions,
): void => {
  setStrokeArrow(requireSpPr(shape), end, options);
  commitAndRefresh(shape);
};

/** Removes any outline override; the shape then inherits stroke from layout. */
export const clearShapeStroke = (shape: SlideShapeData): void => {
  clearStrokeImpl(requireSpPr(shape));
  commitAndRefresh(shape);
};

// ---------------------------------------------------------------------------
// Effects: shadow + glow.

/**
 * Read-back for `setShapeShadow` / `setShapeGlow`. Returns the kind
 * of effect currently on the shape's `<a:effectLst>`, or `null` when
 * none. Decodes the configured color + numeric parameters when
 * present.
 */
export type ShapeEffect =
  | {
      readonly kind: 'shadow';
      readonly color: string;
      readonly blurEmu: number;
      readonly offsetEmu: number;
      readonly angleDeg: number;
      readonly opacity?: number;
    }
  | {
      readonly kind: 'glow';
      readonly color: string;
      readonly radiusEmu: number;
    };

/**
 * Discriminated union covering every effect in
 * `CT_EffectStyleItem` (ECMA-376 §20.1.8.x) — outer shadow, inner
 * shadow, glow, reflection, soft-edge, blur. Returned in document
 * order so renderers can chain filters with the same composition
 * PowerPoint applies.
 *
 * Lengths are EMU; angles are degrees clockwise from 3 o'clock;
 * opacity is a unit fraction (0..1) when the spec exposes one.
 */
export type ShapeEffectAny =
  | {
      readonly kind: 'outerShdw';
      readonly color: string;
      readonly opacity?: number;
      readonly blurEmu: number;
      readonly distEmu: number;
      readonly angleDeg: number;
    }
  | {
      readonly kind: 'innerShdw';
      readonly color: string;
      readonly opacity?: number;
      readonly blurEmu: number;
      readonly distEmu: number;
      readonly angleDeg: number;
    }
  | {
      readonly kind: 'glow';
      readonly color: string;
      readonly opacity?: number;
      readonly radiusEmu: number;
    }
  | {
      readonly kind: 'reflection';
      readonly opacity?: number;
      readonly blurEmu: number;
      readonly distEmu: number;
      readonly angleDeg: number;
    }
  | { readonly kind: 'softEdge'; readonly radiusEmu: number }
  | { readonly kind: 'blur'; readonly radiusEmu: number };

export const getShapeEffect = (shape: SlideShapeData): ShapeEffect | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const effectLst = firstChildElement(spPr, qname('a', 'effectLst', NS.dml));
  if (!effectLst) return null;

  const readColor = (host: XmlElement): { color: string; opacity?: number } => {
    const srgb = firstChildElement(host, qname('a', 'srgbClr', NS.dml));
    if (!srgb) return { color: '' };
    const val = getAttrValue(srgb, qname('', 'val', ''));
    const color = val !== null ? `#${val.toUpperCase()}` : '';
    const alpha = firstChildElement(srgb, qname('a', 'alpha', NS.dml));
    if (alpha) {
      const a = getAttrValue(alpha, qname('', 'val', ''));
      if (a !== null) {
        const n = Number.parseInt(a, 10);
        if (Number.isFinite(n)) return { color, opacity: n / 100000 };
      }
    }
    return { color };
  };

  const outerShdw = firstChildElement(effectLst, qname('a', 'outerShdw', NS.dml));
  if (outerShdw) {
    const blur = Number.parseInt(getAttrValue(outerShdw, qname('', 'blurRad', '')) ?? '0', 10);
    const dist = Number.parseInt(getAttrValue(outerShdw, qname('', 'dist', '')) ?? '0', 10);
    const dirRaw = Number.parseInt(getAttrValue(outerShdw, qname('', 'dir', '')) ?? '0', 10);
    const c = readColor(outerShdw);
    return {
      kind: 'shadow',
      color: c.color,
      blurEmu: blur,
      offsetEmu: dist,
      angleDeg: dirRaw / 60000,
      ...(c.opacity !== undefined ? { opacity: c.opacity } : {}),
    };
  }
  const glow = firstChildElement(effectLst, qname('a', 'glow', NS.dml));
  if (glow) {
    const rad = Number.parseInt(getAttrValue(glow, qname('', 'rad', '')) ?? '0', 10);
    const c = readColor(glow);
    return { kind: 'glow', color: c.color, radiusEmu: rad };
  }
  return null;
};

/**
 * Returns every effect attached to the shape's `<a:effectLst>` in
 * document order — outer shadow, inner shadow, glow, reflection,
 * soft edge, blur. Empty array when no effects apply.
 *
 * Companion to `getShapeEffect`, which is the v1 "first effect only"
 * helper. `getShapeEffects` is what renderers want because PowerPoint
 * composes multiple effects in a single filter (shadow + glow, etc.).
 */
// Parses an `<a:effectLst>` element into the typed effect union.
// Pulled out of `getShapeEffects` so the cascade-aware variant can
// reuse it.
const parseEffectLst = (
  effectLst: XmlElement,
  theme: PresentationTheme | null,
): ShapeEffectAny[] => {
  const readEffectColor = (host: XmlElement): { color: string; opacity?: number } => {
    let inner: XmlElement | null = null;
    for (const c of host.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      if (
        c.name.localName === 'srgbClr' ||
        c.name.localName === 'schemeClr' ||
        c.name.localName === 'sysClr' ||
        c.name.localName === 'prstClr'
      ) {
        inner = c;
        break;
      }
    }
    if (!inner) return { color: '' };
    let opacity: number | undefined;
    const alphaEl = firstChildElement(inner, qname('a', 'alpha', NS.dml));
    if (alphaEl) {
      const a = getAttrValue(alphaEl, qname('', 'val', ''));
      if (a !== null) {
        let n = Number.parseFloat(a);
        if (Number.isFinite(n)) {
          if (Math.abs(n) > 1) n = n / 100000;
          opacity = n;
        }
      }
    }
    const hex = resolveDrawingColor(inner, theme);
    return { color: hex ?? '', ...(opacity !== undefined ? { opacity } : {}) };
  };

  const out: ShapeEffectAny[] = [];
  for (const child of effectLst.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.dml) continue;
    const local = child.name.localName;
    if (local === 'outerShdw' || local === 'innerShdw') {
      const blur = Number.parseInt(getAttrValue(child, qname('', 'blurRad', '')) ?? '0', 10) || 0;
      const dist = Number.parseInt(getAttrValue(child, qname('', 'dist', '')) ?? '0', 10) || 0;
      const dir = Number.parseInt(getAttrValue(child, qname('', 'dir', '')) ?? '0', 10) || 0;
      const c = readEffectColor(child);
      out.push({
        kind: local,
        color: c.color,
        blurEmu: blur,
        distEmu: dist,
        angleDeg: dir / 60000,
        ...(c.opacity !== undefined ? { opacity: c.opacity } : {}),
      });
    } else if (local === 'glow') {
      const rad = Number.parseInt(getAttrValue(child, qname('', 'rad', '')) ?? '0', 10) || 0;
      const c = readEffectColor(child);
      out.push({
        kind: 'glow',
        color: c.color,
        radiusEmu: rad,
        ...(c.opacity !== undefined ? { opacity: c.opacity } : {}),
      });
    } else if (local === 'reflection') {
      const blur = Number.parseInt(getAttrValue(child, qname('', 'blurRad', '')) ?? '0', 10) || 0;
      const dist = Number.parseInt(getAttrValue(child, qname('', 'dist', '')) ?? '0', 10) || 0;
      const dir = Number.parseInt(getAttrValue(child, qname('', 'dir', '')) ?? '0', 10) || 0;
      const endA = getAttrValue(child, qname('', 'endA', ''));
      let opacity: number | undefined;
      if (endA !== null) {
        let n = Number.parseFloat(endA);
        if (Number.isFinite(n)) {
          if (Math.abs(n) > 1) n = n / 100000;
          opacity = n;
        }
      }
      out.push({
        kind: 'reflection',
        blurEmu: blur,
        distEmu: dist,
        angleDeg: dir / 60000,
        ...(opacity !== undefined ? { opacity } : {}),
      });
    } else if (local === 'softEdge') {
      const rad = Number.parseInt(getAttrValue(child, qname('', 'rad', '')) ?? '0', 10) || 0;
      out.push({ kind: 'softEdge', radiusEmu: rad });
    } else if (local === 'blur') {
      const rad = Number.parseInt(getAttrValue(child, qname('', 'rad', '')) ?? '0', 10) || 0;
      out.push({ kind: 'blur', radiusEmu: rad });
    }
  }
  return out;
};

export const getShapeEffects = (
  pres: PresentationData,
  shape: SlideShapeData,
): readonly ShapeEffectAny[] => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return [];
  const effectLst = firstChildElement(spPr, qname('a', 'effectLst', NS.dml));
  if (!effectLst) return [];
  return parseEffectLst(effectLst, getPresentationTheme(pres));
};

/**
 * Same as `getShapeEffects` but walks the layout → master placeholder
 * cascade when the shape itself has no `<a:effectLst>`. Inherits
 * "all or nothing" — once any layer supplies an effect list, that
 * list is used; layers further down aren't merged in. This matches
 * PowerPoint's behaviour (effect lists override rather than compose).
 */
export const getShapeEffectsEffective = (
  pres: PresentationData,
  shape: SlideShapeData,
): readonly ShapeEffectAny[] => {
  const own = getShapeEffects(pres, shape);
  if (own.length > 0) return own;

  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);
  if (phIdx === null && phType === null) return own;

  const theme = getPresentationTheme(pres);
  const layout = getSlideLayout(shape[SHAPE_SLIDE]);
  if (!layout) return own;

  const findPh = (
    shapes: ReadonlyArray<{
      placeholderIdx: number | null;
      placeholderType: string | null;
      element: XmlElement;
    }>,
  ): XmlElement | null => {
    let match = phIdx !== null ? shapes.find((s) => s.placeholderIdx === phIdx) : undefined;
    if (!match && phType !== null) match = shapes.find((s) => s.placeholderType === phType);
    return match?.element ?? null;
  };

  const readEffectsOn = (el: XmlElement): readonly ShapeEffectAny[] => {
    const spPr = firstChildElement(el, qname('p', 'spPr', NS.pml));
    if (!spPr) return [];
    const eff = firstChildElement(spPr, qname('a', 'effectLst', NS.dml));
    if (!eff) return [];
    return parseEffectLst(eff, theme);
  };

  const layoutPh = findPh(layout[LAYOUT_PART].shapes);
  if (layoutPh) {
    const layoutEffects = readEffectsOn(layoutPh);
    if (layoutEffects.length > 0) return layoutEffects;
  }

  const pkg = pres[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const layoutRels = pkg.getRels(layoutPartName);
  if (!layoutRels) return own;
  const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return own;
  const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
  if (!masterPart) return own;
  const masterRoot = parseXml(decode(masterPart.data)).root;
  const { shapes: masterShapes } = readShapeTreeFromCsldRoot(masterRoot, 'sldMaster');
  const masterPh = findPh(masterShapes);
  if (masterPh) {
    const masterEffects = readEffectsOn(masterPh);
    if (masterEffects.length > 0) return masterEffects;
  }
  return own;
};

/**
 * Sets an outer drop shadow on the shape. Defaults: black, 4pt blur,
 * 3pt offset, 45° (down-right). Pass `opacity` (0–1) to soften the
 * shadow.
 */
export const setShapeShadow = (shape: SlideShapeData, options: ShadowOptions = {}): void => {
  setShadow(requireSpPr(shape), options);
  commitAndRefresh(shape);
};

/**
 * Sets a glow around the shape. The radius is in EMU (default 5pt =
 * 63500). Mutually exclusive with `setShapeShadow` in v1 — calling
 * either replaces the prior `<a:effectLst>` entirely.
 */
export const setShapeGlow = (shape: SlideShapeData, options: GlowOptions): void => {
  setGlow(requireSpPr(shape), options);
  commitAndRefresh(shape);
};

/** Removes any effects (shadow / glow / future presets) from the shape. */
export const clearShapeEffects = (shape: SlideShapeData): void => {
  clearEffectsImpl(requireSpPr(shape));
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
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (txBody === null) {
    throw new Error(`shape "${shape[SHAPE_SNAPSHOT].name}" has no <p:txBody>`);
  }
  setTextBody(txBody, value);
  if (options.bullets !== undefined) {
    applyBulletToAllParagraphs(txBody, options.bullets);
  }
  commitAndRefresh(shape);
};

/**
 * Appends `value` to the shape's existing text on a new line. The
 * shape's existing run / paragraph formatting is preserved by
 * `setTextBody`; the new paragraph inherits the same template.
 *
 * Equivalent to `setShapeText(shape, getShapeText(shape) + '\n' + value)`,
 * minus the leading newline when there was no existing text.
 */
export const appendShapeText = (shape: SlideShapeData, value: string): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'shape') {
    throw new Error(
      `appendShapeText only works on text-bearing shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (txBody === null) {
    throw new Error(`shape "${shape[SHAPE_SNAPSHOT].name}" has no <p:txBody>`);
  }
  const existing = shape[SHAPE_SNAPSHOT].text;
  const combined = existing.length === 0 ? value : `${existing}\n${value}`;
  setTextBody(txBody, combined);
  commitAndRefresh(shape);
};

/**
 * Sets the vertical text anchor on the shape's text body
 * (`<a:bodyPr anchor="..."/>`). Choices map to ECMA-376 tokens:
 *
 *   - `'top'`    → `anchor="t"`
 *   - `'center'` → `anchor="ctr"`
 *   - `'bottom'` → `anchor="b"`
 *
 * The bodyPr is created if absent. Throws for non-text-bearing shape
 * kinds.
 */
export type TextAnchor = 'top' | 'center' | 'bottom';

const NAME_A_BODY_PR = qname('a', 'bodyPr', NS.dml);

/**
 * Word wrap mode on a text body. `'square'` (PowerPoint default for
 * textboxes) wraps lines at the shape's width; `'none'` lets text
 * overflow horizontally.
 */
export type TextWrap = 'none' | 'square';

/** Auto-fit mode on a text body. */
export type TextAutoFit =
  | 'none' // <a:noAutofit/>
  | 'normal' // <a:normAutofit/> — shrink text to fit
  | 'shape'; // <a:spAutoFit/> — resize shape to fit text

const AUTO_FIT_LOCALS = new Set(['noAutofit', 'normAutofit', 'spAutoFit']);

const requireBodyPr = (shape: SlideShapeData): XmlElement => {
  const txBody = requireTxBody(shape);
  let bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (bodyPr === null) {
    bodyPr = elem(NAME_A_BODY_PR);
    txBody.children.unshift(bodyPr);
  }
  return bodyPr;
};

/**
 * Sets the text-body word-wrap mode.
 *
 *   - `'square'` writes `wrap="square"` — PowerPoint default for textboxes.
 *   - `'none'`   writes `wrap="none"`  — text can overflow horizontally.
 *
 * Throws for non-text-bearing shape kinds.
 */
export const setShapeTextWrap = (shape: SlideShapeData, wrap: TextWrap): void => {
  const bodyPr = requireBodyPr(shape);
  const ATTR_WRAP = qname('', 'wrap', '');
  bodyPr.attrs = bodyPr.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'wrap'),
  );
  bodyPr.attrs.push(attr(ATTR_WRAP, wrap));
  commitAndRefresh(shape);
};

/** Reads back the bodyPr `wrap` attribute, or `null` when absent. */
export const getShapeTextWrap = (shape: SlideShapeData): TextWrap | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  const v = getAttrValue(bodyPr, qname('', 'wrap', ''));
  if (v === 'none' || v === 'square') return v;
  return null;
};

/**
 * Sets the text-body auto-fit mode:
 *
 *   - `'none'`   → `<a:noAutofit/>`
 *   - `'normal'` → `<a:normAutofit/>`   shrink text to fit the shape
 *   - `'shape'`  → `<a:spAutoFit/>`     grow the shape to fit text
 *
 * Replaces any prior auto-fit child on `<a:bodyPr>`. Throws for
 * non-text-bearing shape kinds.
 */
export const setShapeTextAutoFit = (shape: SlideShapeData, mode: TextAutoFit): void => {
  const bodyPr = requireBodyPr(shape);
  bodyPr.children = bodyPr.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        AUTO_FIT_LOCALS.has(c.name.localName)
      ),
  );
  const local = mode === 'none' ? 'noAutofit' : mode === 'normal' ? 'normAutofit' : 'spAutoFit';
  bodyPr.children.push(elem(qname('a', local, NS.dml)));
  commitAndRefresh(shape);
};

/**
 * Reads back the bodyPr auto-fit child, or `null` when none is
 * present (PowerPoint applies a layout-inherited default in that case).
 */
export const getShapeTextAutoFit = (shape: SlideShapeData): TextAutoFit | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  for (const c of bodyPr.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (c.name.localName === 'noAutofit') return 'none';
    if (c.name.localName === 'normAutofit') return 'normal';
    if (c.name.localName === 'spAutoFit') return 'shape';
  }
  return null;
};

/**
 * Reads the scale parameters PowerPoint stores on `<a:normAutofit>`
 * once it has shrunk a text body to fit. Returns `null` if the body
 * doesn't carry `<a:normAutofit>` or the attributes are absent. Both
 * fields are unitless ratios in `[0, 1]`:
 *
 *   - `fontScale`     — multiply every run's font size by this. Default `1`.
 *   - `lnSpcReduction` — subtract from the line-height ratio. Default `0`.
 *
 * Companion to `getShapeTextAutoFit`. Renderers that want to match
 * PowerPoint's actual on-screen text size apply these factors to the
 * authored font sizes; without them, every long title overflows.
 */
export const getShapeTextAutoFitParams = (
  shape: SlideShapeData,
): { fontScale: number; lnSpcReduction: number } | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  for (const c of bodyPr.children) {
    if (
      c.kind === 'element' &&
      c.name.namespaceURI === NS.dml &&
      c.name.localName === 'normAutofit'
    ) {
      const fsRaw = getAttrValue(c, qname('', 'fontScale', ''));
      const lsRaw = getAttrValue(c, qname('', 'lnSpcReduction', ''));
      const fs = fsRaw === null ? 100_000 : Number.parseInt(fsRaw, 10);
      const ls = lsRaw === null ? 0 : Number.parseInt(lsRaw, 10);
      return {
        fontScale: Number.isFinite(fs) ? fs / 100_000 : 1,
        lnSpcReduction: Number.isFinite(ls) ? ls / 100_000 : 0,
      };
    }
  }
  return null;
};

/**
 * Reads back the vertical text anchor on the shape's `<a:bodyPr>`.
 * Maps the ECMA-376 tokens back to the public union:
 *
 *   `'t'` → `'top'`, `'ctr'` → `'center'`, `'b'` → `'bottom'`
 *
 * Returns `null` when the bodyPr is absent or has no anchor attribute.
 */
export const getShapeTextAnchor = (shape: SlideShapeData): TextAnchor | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  const v = getAttrValue(bodyPr, qname('', 'anchor', ''));
  if (v === 't') return 'top';
  if (v === 'ctr') return 'center';
  if (v === 'b') return 'bottom';
  return null;
};

/**
 * Reads back the internal margins of the shape's text frame. Sides
 * that are absent in the XML default to `null` (PowerPoint applies
 * its built-in default for the missing side).
 */
/**
 * Reads the multi-column layout on a text body — `<a:bodyPr
 * numCol="N" spcCol="EMU"/>`. Returns `null` when columns aren't
 * configured (the default single column). `gapEmu` is the
 * inter-column gap in EMU; omitted when `<a:bodyPr>` has no
 * `spcCol` attribute.
 */
export const getShapeTextColumns = (
  shape: SlideShapeData,
): { count: number; gapEmu?: number } | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  const numColRaw = getAttrValue(bodyPr, qname('', 'numCol', ''));
  if (numColRaw === null) return null;
  const count = Number.parseInt(numColRaw, 10);
  if (!Number.isFinite(count) || count < 2) return null;
  const gapRaw = getAttrValue(bodyPr, qname('', 'spcCol', ''));
  if (gapRaw !== null) {
    const g = Number.parseInt(gapRaw, 10);
    if (Number.isFinite(g)) return { count, gapEmu: g };
  }
  return { count };
};

/**
 * Reads the shape's text-body rotation from `<a:bodyPr rot="N"/>`.
 * `rot` is stored in 60000ths of a degree (OOXML angle units); the
 * returned value is in degrees. Positive values rotate clockwise per
 * PowerPoint's convention. Returns `null` when the attribute is
 * absent.
 *
 * Distinct from the shape's own `<p:xfrm rot=…>` (the geometry
 * rotation surfaced via the shape's `rotation`); `bodyPr rot` rotates
 * the text body inside the shape without rotating the shape itself.
 */
export const getShapeTextBodyRotationDeg = (shape: SlideShapeData): number | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  const v = getAttrValue(bodyPr, qname('', 'rot', ''));
  if (v === null) return null;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n / 60000 : null;
};

/**
 * Sets the shape's text-body rotation (`<a:bodyPr rot="N"/>`), measured
 * in degrees. Positive rotates clockwise per PowerPoint's convention.
 * Passing `null` clears the attribute so the shape inherits the default
 * (`0`). Throws for non-text-bearing shape kinds.
 *
 * Companion to `setShapeRotation`, which rotates the *whole* shape
 * via `<p:xfrm rot>`. `bodyPr rot` rotates only the text inside.
 */
export const setShapeTextBodyRotationDeg = (
  shape: SlideShapeData,
  rotationDeg: number | null,
): void => {
  const bodyPr = requireBodyPr(shape);
  // Strip any prior rot attribute.
  bodyPr.attrs = bodyPr.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'rot'),
  );
  if (rotationDeg !== null && rotationDeg !== 0) {
    bodyPr.attrs.push(attr(qname('', 'rot', ''), String(Math.round(rotationDeg * 60000))));
  }
  commitAndRefresh(shape);
};

/**
 * Reads the shape's text-direction token from `<a:bodyPr vert="…"/>`.
 * Per ECMA-376 §17.18.93 `ST_TextVerticalType`:
 *
 *   - `horz` — default left-to-right, top-to-bottom (returns `null`).
 *   - `vert` — 90° rotation, lines run top-to-bottom, columns right-to-left.
 *   - `vert270` — 270° rotation, lines top-to-bottom, columns left-to-right.
 *   - `wordArtVert` — characters not rotated, stacked vertically.
 *   - `eaVert` — East-Asian vertical: characters upright, columns right-to-left.
 *   - `mongolianVert` — Mongolian: rotated 90°, columns left-to-right.
 *   - `wordArtVertRtl` — RTL word-art stacked vertically.
 *
 * Returns `null` when the attribute is absent or set to the default
 * `horz`.
 */
export const getShapeTextDirection = (
  shape: SlideShapeData,
): 'vert' | 'vert270' | 'wordArtVert' | 'eaVert' | 'mongolianVert' | 'wordArtVertRtl' | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  const v = getAttrValue(bodyPr, qname('', 'vert', ''));
  if (
    v === 'vert' ||
    v === 'vert270' ||
    v === 'wordArtVert' ||
    v === 'eaVert' ||
    v === 'mongolianVert' ||
    v === 'wordArtVertRtl'
  )
    return v;
  return null;
};

export const getShapeTextMargins = (
  shape: SlideShapeData,
): {
  readonly left: number | null;
  readonly top: number | null;
  readonly right: number | null;
  readonly bottom: number | null;
} | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  const readSide = (local: string): number | null => {
    const v = getAttrValue(bodyPr, qname('', local, ''));
    if (v === null) return null;
    const n = Number.parseInt(v, 10);
    return Number.isFinite(n) ? n : null;
  };
  return {
    left: readSide('lIns'),
    top: readSide('tIns'),
    right: readSide('rIns'),
    bottom: readSide('bIns'),
  };
};

/**
 * Resolves the effective `<a:bodyPr>` properties — anchor, wrap, vertical
 * direction, and inset margins — by walking the layout / master cascade
 * the same way `getShapeRunFormatEffective` walks rPr. Returns the
 * innermost value that the cascade supplies, or `null` for properties
 * neither the shape nor any inherited placeholder authors.
 *
 * Companion to `getShapeTextAnchor` / `getShapeTextWrap` /
 * `getShapeTextDirection` / `getShapeTextMargins`, which only report the
 * literal value on the shape itself.
 */
export const getShapeBodyPrEffective = (
  pres: PresentationData,
  shape: SlideShapeData,
): {
  anchor: TextAnchor | null;
  wrap: TextWrap | null;
  vert: ReturnType<typeof getShapeTextDirection>;
  margins: { left: number | null; top: number | null; right: number | null; bottom: number | null };
} => {
  const result = {
    anchor: null as TextAnchor | null,
    wrap: null as TextWrap | null,
    vert: null as ReturnType<typeof getShapeTextDirection>,
    margins: {
      left: null as number | null,
      top: null as number | null,
      right: null as number | null,
      bottom: null as number | null,
    },
  };
  const parseBodyPr = (bodyPr: XmlElement): void => {
    if (result.anchor === null) {
      const a = getAttrValue(bodyPr, qname('', 'anchor', ''));
      if (a === 't') result.anchor = 'top';
      else if (a === 'ctr') result.anchor = 'center';
      else if (a === 'b') result.anchor = 'bottom';
    }
    if (result.wrap === null) {
      const w = getAttrValue(bodyPr, qname('', 'wrap', ''));
      if (w === 'square') result.wrap = 'square';
      else if (w === 'none') result.wrap = 'none';
    }
    if (result.vert === null) {
      const v = getAttrValue(bodyPr, qname('', 'vert', ''));
      if (
        v === 'vert' ||
        v === 'vert270' ||
        v === 'wordArtVert' ||
        v === 'eaVert' ||
        v === 'mongolianVert' ||
        v === 'wordArtVertRtl'
      )
        result.vert = v;
    }
    for (const side of ['l', 't', 'r', 'b'] as const) {
      const target =
        side === 'l' ? 'left' : side === 't' ? 'top' : side === 'r' ? 'right' : 'bottom';
      if (result.margins[target] !== null) continue;
      const v = getAttrValue(bodyPr, qname('', `${side}Ins`, ''));
      if (v === null) continue;
      const n = Number.parseInt(v, 10);
      if (Number.isFinite(n)) result.margins[target] = n;
    }
  };

  // 1. The shape's own bodyPr.
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (txBody) {
    const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
    if (bodyPr) parseBodyPr(bodyPr);
  }

  // 2-3. Walk layout placeholder and master placeholder bodyPr.
  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);
  const slide = shape[SHAPE_SLIDE];
  const layout = getSlideLayout(slide);
  if (!layout) return result;

  const findPh = (
    shapes: ReadonlyArray<{
      placeholderIdx: number | null;
      placeholderType: string | null;
      element: XmlElement;
    }>,
  ): XmlElement | null => {
    let match = phIdx !== null ? shapes.find((s) => s.placeholderIdx === phIdx) : undefined;
    if (!match && phType !== null) {
      match = shapes.find((s) => s.placeholderType === phType);
    }
    return match?.element ?? null;
  };

  const layoutPhEl = findPh(layout[LAYOUT_PART].shapes);
  if (layoutPhEl) {
    const layoutTxBody = firstChildElement(layoutPhEl, NAME_TX_BODY);
    if (layoutTxBody) {
      const bodyPr = firstChildElement(layoutTxBody, NAME_A_BODY_PR);
      if (bodyPr) parseBodyPr(bodyPr);
    }
  }

  const pkg = pres[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const layoutRels = pkg.getRels(layoutPartName);
  if (!layoutRels) return result;
  const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return result;
  const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
  if (!masterPart) return result;
  const masterRoot = parseXml(decode(masterPart.data)).root;
  const { shapes: masterShapes } = readShapeTreeFromCsldRoot(masterRoot, 'sldMaster');
  const masterPhEl = findPh(masterShapes);
  if (masterPhEl) {
    const masterTxBody = firstChildElement(masterPhEl, NAME_TX_BODY);
    if (masterTxBody) {
      const bodyPr = firstChildElement(masterTxBody, NAME_A_BODY_PR);
      if (bodyPr) parseBodyPr(bodyPr);
    }
  }
  return result;
};

export const setShapeTextAnchor = (shape: SlideShapeData, anchor: TextAnchor): void => {
  const txBody = requireTxBody(shape);
  let bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (bodyPr === null) {
    bodyPr = elem(NAME_A_BODY_PR);
    txBody.children.unshift(bodyPr);
  }
  const token = anchor === 'top' ? 't' : anchor === 'center' ? 'ctr' : 'b';
  const ATTR_ANCHOR = qname('', 'anchor', '');
  // Replace any existing anchor attribute.
  bodyPr.attrs = bodyPr.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'anchor'),
  );
  bodyPr.attrs.push(attr(ATTR_ANCHOR, token));
  commitAndRefresh(shape);
};

/**
 * Sets the internal margins of the shape's text frame in EMU. Each
 * side is independent; omitted sides keep their current value (or the
 * layout-inherited default when the attribute is absent).
 *
 * PowerPoint's defaults for a textbox: left/right 91440 (0.1in),
 * top/bottom 45720 (0.05in).
 *
 *   setShapeTextMargins(shape, { left: 0, right: 0 }); // flush-left text
 */
export const setShapeTextMargins = (
  shape: SlideShapeData,
  margins: { left?: number; top?: number; right?: number; bottom?: number },
): void => {
  const txBody = requireTxBody(shape);
  let bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (bodyPr === null) {
    bodyPr = elem(NAME_A_BODY_PR);
    txBody.children.unshift(bodyPr);
  }
  const writes: Array<{ name: string; value: number }> = [];
  if (margins.left !== undefined) writes.push({ name: 'lIns', value: margins.left });
  if (margins.top !== undefined) writes.push({ name: 'tIns', value: margins.top });
  if (margins.right !== undefined) writes.push({ name: 'rIns', value: margins.right });
  if (margins.bottom !== undefined) writes.push({ name: 'bIns', value: margins.bottom });

  const localsToClear = new Set(writes.map((w) => w.name));
  bodyPr.attrs = bodyPr.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && localsToClear.has(a.name.localName)),
  );
  for (const w of writes) {
    bodyPr.attrs.push(attr(qname('', w.name, ''), String(Math.round(w.value))));
  }
  commitAndRefresh(shape);
};

/** Sets the bullet style on every paragraph in the shape's text body. */
export const setShapeBullets = (shape: SlideShapeData, style: BulletStyle): void => {
  applyBulletToAllParagraphs(requireTxBody(shape), style);
  commitAndRefresh(shape);
};

/** Sets the horizontal alignment of every paragraph in the shape's text. */
export const setShapeAlignment = (shape: SlideShapeData, align: ParagraphAlignment): void => {
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
 * One inline element in a paragraph as ordered: a literal text run
 * (`<a:r>`), a field substitution (`<a:fld>` — slide number, date, etc.),
 * or a line break (`<a:br>`). Renderers walk this list instead of the
 * strict `<a:r>`-only `getShapeRunCount` / `getShapeRunText` pair when
 * they need to reproduce the paragraph's full visible content.
 *
 * `text` is the cached value (`<a:t>` content for `r` and `fld`; `''`
 * for `br`). `format` is the literal `<a:rPr>` on the element when
 * present; use `getShapeRunFormatEffective` to walk inheritance.
 *
 * Field kinds (`fld.type`): typical ECMA-376 `ST_TextFieldType` tokens
 * are `slidenum`, `datetime` (variants `1`..`13`), `presentationDate`,
 * `headerfooter`, `footer`, etc. Unrecognised tokens come through
 * unchanged so renderers can decide whether to substitute live values.
 */
export type ShapeParagraphElement =
  | { readonly kind: 'r'; readonly text: string; readonly format: TextFormat | null }
  | {
      readonly kind: 'fld';
      readonly text: string;
      readonly format: TextFormat | null;
      readonly type: string | null;
    }
  | { readonly kind: 'br'; readonly format: TextFormat | null };

/**
 * Returns the inline children of a paragraph in document order — runs,
 * field placeholders, and line breaks. Used by renderers that need to
 * reproduce the paragraph faithfully (the `<a:r>`-only run accessors
 * silently drop fields and breaks).
 */
export const getShapeParagraphElements = (
  shape: SlideShapeData,
  paragraphIndex: number,
): ReadonlyArray<ShapeParagraphElement> => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const out: ShapeParagraphElement[] = [];
  const readT = (parent: XmlElement): string => {
    const tEl = firstChildElement(parent, NAME_A_T);
    if (!tEl) return '';
    let acc = '';
    for (const c of tEl.children) {
      if (c.kind === 'text' || c.kind === 'cdata') acc += c.data;
    }
    return acc;
  };
  const readFmt = (parent: XmlElement): TextFormat | null => {
    const rPr = firstChildElement(parent, NAME_A_RPR);
    if (!rPr) return null;
    return parseRPrLikeElement(rPr) as TextFormat;
  };
  for (const child of paragraph.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.dml) continue;
    if (child.name.localName === 'r') {
      out.push({ kind: 'r', text: readT(child), format: readFmt(child) });
    } else if (child.name.localName === 'fld') {
      const type = getAttrValue(child, qname('', 'type', ''));
      out.push({ kind: 'fld', text: readT(child), format: readFmt(child), type });
    } else if (child.name.localName === 'br') {
      out.push({ kind: 'br', format: readFmt(child) });
    }
  }
  return out;
};

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
 * Sets `<a:hlinkClick>` on a single run. Per-run counterpart to
 * `setShapeHyperlink` (which targets every run in the shape). Pass
 * `null` to clear the link on that run alone — other runs are
 * untouched. Allocates or reuses a hyperlink rel on the slide
 * exactly like the shape-level setter.
 */
export const setShapeRunHyperlink = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
  url: string | null,
): void => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  let rPr = firstChildElement(run, qname('a', 'rPr', NS.dml));
  if (rPr === null) {
    rPr = elem(qname('a', 'rPr', NS.dml));
    run.children.unshift(rPr);
  }
  rPr.children = rPr.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        c.name.localName === 'hlinkClick'
      ),
  );
  if (url !== null) {
    const slide = shape[SHAPE_SLIDE];
    const pkg = slide[INTERNAL_PACKAGE];
    const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
    const existing = rels.items.find(
      (r) => r.type === REL_TYPES.hyperlink && r.target === url && r.targetMode === 'External',
    );
    let rId: string;
    if (existing) {
      rId = existing.id;
    } else {
      rId = nextRelId(rels.items.map((r) => r.id));
      rels.items.push({
        id: rId,
        type: REL_TYPES.hyperlink,
        target: url,
        targetMode: 'External',
      });
      pkg.setRels(slide[SLIDE_PART_NAME], rels);
    }
    rPr.children.push(
      elem(qname('a', 'hlinkClick', NS.dml), {
        attrs: [attr(qname('r', 'id', NS.officeDocRels), rId)],
      }),
    );
  }
  commitAndRefresh(shape);
};

/**
 * Reads the external URL on a single run's `<a:hlinkClick>`. Per-run
 * counterpart to `getShapeHyperlink` (which only surfaces the first
 * link it finds). Returns `null` when this run has no link, or the
 * link's `r:id` resolves to a non-hyperlink rel.
 */
export const getShapeRunHyperlink = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): string | null => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  const rPr = firstChildElement(run, qname('a', 'rPr', NS.dml));
  if (!rPr) return null;
  const hlink = firstChildElement(rPr, qname('a', 'hlinkClick', NS.dml));
  if (!hlink) return null;
  const rId = getAttrValue(hlink, qname('r', 'id', NS.officeDocRels));
  if (!rId) return null;
  const slide = shape[SHAPE_SLIDE];
  const rels = slide[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;
  const rel = rels.items.find((x) => x.id === rId);
  if (rel?.type === REL_TYPES.hyperlink && rel.targetMode === 'External') return rel.target;
  return null;
};

/**
 * Reads the tooltip text on the shape's `<a:hlinkClick tooltip="…"/>`.
 * Returns `null` when no hyperlink is set or the link doesn't author
 * a tooltip. Tooltips show up in PowerPoint when the user hovers over
 * a linked shape in slide-show mode.
 */
export const getShapeHyperlinkTooltip = (shape: SlideShapeData): string | null => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) return null;
  const hlink = firstChildElement(cNvPr, NAME_HLINK_CLICK_FN);
  if (!hlink) return null;
  const tt = getAttrValue(hlink, qname('', 'tooltip', ''));
  return tt ?? null;
};

/**
 * Reads the tooltip on a per-run `<a:rPr><a:hlinkClick tooltip="…"/>`.
 * Same semantics as `getShapeHyperlinkTooltip` but scoped to a single
 * run.
 */
export const getShapeRunHyperlinkTooltip = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): string | null => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  const rPr = firstChildElement(run, qname('a', 'rPr', NS.dml));
  if (!rPr) return null;
  const hlink = firstChildElement(rPr, qname('a', 'hlinkClick', NS.dml));
  if (!hlink) return null;
  const tt = getAttrValue(hlink, qname('', 'tooltip', ''));
  return tt ?? null;
};

/**
 * Same as `getShapeClickAction` but reads the per-run
 * `<a:rPr><a:hlinkClick action=… r:id=…/>`. Recognises:
 *
 *   - `{ kind: 'url', url }` — external hyperlink rel
 *   - `{ kind: 'slide', slide }` — slide-jump action + slide rel
 *   - `{ kind: 'nextSlide' | 'prevSlide' | 'firstSlide' | 'lastSlide' }`
 *
 * Returns `null` for runs without an action or unknown action tokens.
 */
export const getShapeRunClickAction = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): ShapeClickAction | null => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  const rPr = firstChildElement(run, qname('a', 'rPr', NS.dml));
  if (!rPr) return null;
  const hlink = firstChildElement(rPr, qname('a', 'hlinkClick', NS.dml));
  if (!hlink) return null;
  const action = getAttrValue(hlink, qname('', 'action', ''));
  const rId = getAttrValue(hlink, qname('r', 'id', NS.officeDocRels));

  if (action === 'ppaction://hlinkshowjump?jump=nextslide') return { kind: 'nextSlide' };
  if (action === 'ppaction://hlinkshowjump?jump=previousslide') return { kind: 'prevSlide' };
  if (action === 'ppaction://hlinkshowjump?jump=firstslide') return { kind: 'firstSlide' };
  if (action === 'ppaction://hlinkshowjump?jump=lastslide') return { kind: 'lastSlide' };

  if (rId === null || rId === '') return null;
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;
  const rel = rels.items.find((r) => r.id === rId);
  if (!rel) return null;
  if (action === 'ppaction://hlinksldjump' && rel.type === REL_TYPES.slide) {
    const targetPartName = rel.target.startsWith('/')
      ? partName(rel.target)
      : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
    const pres: PresentationData = { [INTERNAL_PACKAGE]: pkg, _slidesCache: null };
    for (const candidate of getSlides(pres)) {
      if (candidate[SLIDE_PART_NAME] === targetPartName) return { kind: 'slide', slide: candidate };
    }
    return null;
  }
  if (rel.type === REL_TYPES.hyperlink && rel.targetMode === 'External') {
    return { kind: 'url', url: rel.target };
  }
  return null;
};

const NAME_A_PPR = qname('a', 'pPr', NS.dml);
const ATTR_LVL = qname('', 'lvl', '');
const ATTR_ALGN_FN = qname('', 'algn', '');

const ensurePPr = (paragraph: XmlElement): XmlElement => {
  const existing = firstChildElement(paragraph, NAME_A_PPR);
  if (existing !== null) return existing;
  const fresh = elem(NAME_A_PPR);
  // <a:pPr> must be the first child of <a:p>.
  paragraph.children.unshift(fresh);
  return fresh;
};

const alignTokenForFn = (a: ParagraphAlignment): string => {
  switch (a) {
    case 'left':
    case 'l':
      return 'l';
    case 'center':
    case 'ctr':
      return 'ctr';
    case 'right':
    case 'r':
      return 'r';
    case 'justify':
    case 'just':
      return 'just';
    case 'distribute':
    case 'dist':
      return 'dist';
    default:
      return a;
  }
};

/**
 * Sets the horizontal alignment of a single paragraph. Same token set
 * as `setShapeAlignment`. Other paragraphs are untouched.
 */
export const setParagraphAlignment = (
  shape: SlideShapeData,
  paragraphIndex: number,
  align: ParagraphAlignment,
): void => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = ensurePPr(paragraph);
  pPr.attrs = pPr.attrs.filter((a) => a.name.localName !== 'algn');
  pPr.attrs.push(attr(ATTR_ALGN_FN, alignTokenForFn(align)));
  commitAndRefresh(shape);
};

/**
 * Sets the paragraph's nesting level (`<a:pPr lvl="N"/>`). Levels are
 * 0-indexed; PowerPoint accepts 0 through 8. Pass `0` to clear an
 * existing level — `<a:pPr lvl="0"/>` is the same as omitting the attr.
 *
 * Used in tandem with bullets to author nested lists:
 *
 *   setShapeText(shape, 'Item 1\nNested\nItem 2');
 *   setShapeBullets(shape, 'bullet');
 *   setParagraphLevel(shape, 1, 1);  // indent the second line
 */
export const setParagraphLevel = (
  shape: SlideShapeData,
  paragraphIndex: number,
  level: number,
): void => {
  if (!Number.isInteger(level) || level < 0 || level > 8) {
    throw new RangeError(`paragraph level must be an integer in [0, 8], got ${level}`);
  }
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = ensurePPr(paragraph);
  pPr.attrs = pPr.attrs.filter((a) => a.name.localName !== 'lvl');
  if (level > 0) pPr.attrs.push(attr(ATTR_LVL, String(level)));
  commitAndRefresh(shape);
};

/**
 * Reads the paragraph's horizontal alignment. Returns `null` when no
 * `algn` attribute is present (inherits from layout / master).
 */
export const getParagraphAlignment = (
  shape: SlideShapeData,
  paragraphIndex: number,
): ParagraphAlignment | null => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (pPr === null) return null;
  const v = getAttrValue(pPr, ATTR_ALGN_FN);
  return (v as ParagraphAlignment | null) ?? null;
};

/**
 * Reads the paragraph's nesting level (`lvl` attribute), or `0` when
 * absent — PowerPoint's default. Returns `null` for non-existent
 * paragraphs.
 */
export const getParagraphLevel = (shape: SlideShapeData, paragraphIndex: number): number => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (pPr === null) return 0;
  const v = getAttrValue(pPr, ATTR_LVL);
  if (v === null) return 0;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n : 0;
};

/**
 * Sets the spacing before and/or after a paragraph, in points (where
 * a "point" is 1/72 inch). PowerPoint stores these as hundredths of a
 * point inside `<a:pPr><a:spcBef>/<a:spcAft><a:spcPts val="…"/>` —
 * the helper converts.
 *
 *   setParagraphSpacing(shape, 0, { beforePts: 6, afterPts: 3 });
 *
 * Omitting a side keeps the existing value (or layout default).
 * Passing a side as `null` removes that spacing element.
 */
export const setParagraphSpacing = (
  shape: SlideShapeData,
  paragraphIndex: number,
  opts: { beforePts?: number | null; afterPts?: number | null },
): void => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = ensurePPr(paragraph);

  const writeSide = (localName: 'spcBef' | 'spcAft', value: number | null | undefined): void => {
    if (value === undefined) return;
    pPr.children = pPr.children.filter(
      (c) =>
        !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === localName),
    );
    if (value === null) return;
    if (!Number.isFinite(value) || value < 0) {
      throw new RangeError(`paragraph ${localName} must be a non-negative number, got ${value}`);
    }
    const spcEl = elem(qname('a', localName, NS.dml), {
      children: [
        elem(qname('a', 'spcPts', NS.dml), {
          attrs: [attr(qname('', 'val', ''), String(Math.round(value * 100)))],
        }),
      ],
    });
    pPr.children.push(spcEl);
  };

  writeSide('spcBef', opts.beforePts);
  writeSide('spcAft', opts.afterPts);
  commitAndRefresh(shape);
};

/**
 * Reads back paragraph spacing in points. Returns `{ beforePts,
 * afterPts }`; each side is `null` when no `<a:spcBef>` / `<a:spcAft>`
 * is present or when the inner element isn't `<a:spcPts>` (percentage
 * spacing is reported as `null` for now).
 */
export const getParagraphSpacing = (
  shape: SlideShapeData,
  paragraphIndex: number,
): { readonly beforePts: number | null; readonly afterPts: number | null } => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (!pPr) return { beforePts: null, afterPts: null };
  const readSide = (localName: 'spcBef' | 'spcAft'): number | null => {
    const side = firstChildElement(pPr, qname('a', localName, NS.dml));
    if (!side) return null;
    const spcPts = firstChildElement(side, qname('a', 'spcPts', NS.dml));
    if (!spcPts) return null;
    const v = getAttrValue(spcPts, qname('', 'val', ''));
    if (v === null) return null;
    const n = Number.parseInt(v, 10);
    return Number.isFinite(n) ? n / 100 : null;
  };
  return { beforePts: readSide('spcBef'), afterPts: readSide('spcAft') };
};

/**
 * Reads the paragraph's left / right / first-line indents from
 * `<a:pPr marL="…" marR="…" indent="…"/>`. Each is in EMU (matching
 * PowerPoint's internal storage); positive means a positive indent,
 * negative `indent` is a hanging indent (typical for bullets).
 *
 * Returns `null` for sides the paragraph doesn't set (those inherit
 * from the layout / master).
 */
export const getParagraphIndent = (
  shape: SlideShapeData,
  paragraphIndex: number,
): { leftEmu: number | null; rightEmu: number | null; firstLineEmu: number | null } => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (!pPr) return { leftEmu: null, rightEmu: null, firstLineEmu: null };
  const read = (name: string): number | null => {
    const raw = getAttrValue(pPr, qname('', name, ''));
    if (raw === null) return null;
    const n = Number.parseInt(raw, 10);
    return Number.isFinite(n) ? n : null;
  };
  return {
    leftEmu: read('marL'),
    rightEmu: read('marR'),
    firstLineEmu: read('indent'),
  };
};

/**
 * Reads the paragraph's `<a:lnSpc>` line spacing. PowerPoint stores
 * line spacing two ways:
 *
 *   - Multiple of the natural line height — `<a:spcPct val="150000"/>`
 *     (= 1.5×). Returns `{ kind: 'pct', value }` with value as the unit
 *     fraction (1.5).
 *   - Fixed points — `<a:spcPts val="2400"/>` (= 24pt). Returns
 *     `{ kind: 'pts', value }` with value in points.
 *
 * Returns `null` when no `<a:lnSpc>` is present (the paragraph
 * inherits line spacing from the layout / master).
 */
export const getParagraphLineSpacing = (
  shape: SlideShapeData,
  paragraphIndex: number,
):
  | { readonly kind: 'pct'; readonly value: number }
  | { readonly kind: 'pts'; readonly value: number }
  | null => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (!pPr) return null;
  const lnSpc = firstChildElement(pPr, qname('a', 'lnSpc', NS.dml));
  if (!lnSpc) return null;
  const pct = firstChildElement(lnSpc, qname('a', 'spcPct', NS.dml));
  if (pct) {
    const v = getAttrValue(pct, qname('', 'val', ''));
    if (v !== null) {
      let n = Number.parseFloat(v);
      if (Number.isFinite(n)) {
        if (Math.abs(n) > 1) n = n / 100000;
        return { kind: 'pct', value: n };
      }
    }
  }
  const pts = firstChildElement(lnSpc, qname('a', 'spcPts', NS.dml));
  if (pts) {
    const v = getAttrValue(pts, qname('', 'val', ''));
    if (v !== null) {
      const n = Number.parseInt(v, 10);
      if (Number.isFinite(n)) return { kind: 'pts', value: n / 100 };
    }
  }
  return null;
};

/**
 * Reads back the bullet style on a single paragraph, or `null` when
 * no `<a:buChar>` / `<a:buAutoNum>` / `<a:buNone>` is present (the
 * paragraph inherits its bullet from the layout / master).
 */
export const getParagraphBullet = (
  shape: SlideShapeData,
  paragraphIndex: number,
): BulletStyle | null => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (pPr === null) return null;
  for (const c of pPr.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (c.name.localName === 'buNone') return 'none';
    if (c.name.localName === 'buChar') {
      const char = getAttrValue(c, qname('', 'char', ''));
      if (char === '•') return 'bullet';
      if (char !== null) return { char };
    }
    if (c.name.localName === 'buAutoNum') {
      const t = getAttrValue(c, qname('', 'type', ''));
      if (t === 'arabicPeriod') return 'number';
      if (t !== null) return { autoNum: t };
    }
  }
  return null;
};

/**
 * Returns `true` when the paragraph uses an image as its bullet
 * (`<a:pPr><a:buBlip r:embed="…"/>`). Renderers without image
 * support should fall back to a generic bullet glyph.
 *
 * The underlying rId / image bytes aren't surfaced here — resolving
 * that would need the rels of the layout / master the paragraph
 * inherits from, which can be cumbersome. Knowing that the bullet
 * *is* an image is usually enough for the UI to pick a fallback.
 */
export const isParagraphBulletPicture = (
  shape: SlideShapeData,
  paragraphIndex: number,
): boolean => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (!pPr) return false;
  return firstChildElement(pPr, qname('a', 'buBlip', NS.dml)) !== null;
};

/**
 * Reads the bullet's per-paragraph color, size, and font overrides —
 * `<a:buClr>` (theme-resolved hex), `<a:buSzPct>` / `<a:buSzPts>`
 * (size relative to run or fixed pt), and `<a:buFont typeface="…"/>`.
 *
 * Returns `{ color: null, sizePct: null, sizePts: null, font: null }`
 * when the paragraph doesn't override any of them (the bullet inherits
 * from the run / layout).
 */
export const getParagraphBulletStyle = (
  pres: PresentationData,
  shape: SlideShapeData,
  paragraphIndex: number,
): {
  color: string | null;
  sizePct: number | null;
  sizePts: number | null;
  font: string | null;
} => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (!pPr) return { color: null, sizePct: null, sizePts: null, font: null };
  const theme = getPresentationTheme(pres);
  let color: string | null = null;
  let sizePct: number | null = null;
  let sizePts: number | null = null;
  let font: string | null = null;
  const buClr = firstChildElement(pPr, qname('a', 'buClr', NS.dml));
  if (buClr) {
    for (const c of buClr.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      color = resolveDrawingColor(c, theme);
      break;
    }
  }
  const buSzPct = firstChildElement(pPr, qname('a', 'buSzPct', NS.dml));
  if (buSzPct) {
    const v = getAttrValue(buSzPct, qname('', 'val', ''));
    if (v !== null) {
      let n = Number.parseFloat(v);
      if (Number.isFinite(n)) {
        if (Math.abs(n) > 1) n = n / 100000;
        sizePct = n;
      }
    }
  }
  const buSzPts = firstChildElement(pPr, qname('a', 'buSzPts', NS.dml));
  if (buSzPts) {
    const v = getAttrValue(buSzPts, qname('', 'val', ''));
    if (v !== null) {
      const n = Number.parseInt(v, 10);
      if (Number.isFinite(n)) sizePts = n / 100;
    }
  }
  const buFont = firstChildElement(pPr, qname('a', 'buFont', NS.dml));
  if (buFont) {
    const t = getAttrValue(buFont, qname('', 'typeface', ''));
    if (t !== null) font = t;
  }
  return { color, sizePct, sizePts, font };
};

/**
 * Sets the bullet style on a single paragraph. Same `BulletStyle` shape
 * as `setShapeBullets` — pass `'bullet'` / `'number'` / `'none'` or an
 * object like `{ char: '◆' }` / `{ autoNum: 'romanLcPeriod' }`.
 */
export const setParagraphBullet = (
  shape: SlideShapeData,
  paragraphIndex: number,
  style: BulletStyle,
): void => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  applyBulletToParagraph(paragraph, style);
  commitAndRefresh(shape);
};

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

// -- Color transforms (ECMA-376 §20.1.2.3.x) --------------------------------
//
// DrawingML color elements (`<a:srgbClr>`, `<a:schemeClr>`, `<a:sysClr>`,
// `<a:prstClr>`) may carry one or more transform children — `lumMod`,
// `lumOff`, `shade`, `tint`, `satMod`, `hueMod`, `alpha`, `gray`, `inv`,
// `comp`, etc. — that adjust the base color before it's painted. Real
// templates use them heavily for "tinted accent" backgrounds and "shaded
// hover" states, so any visual-fidelity story has to apply them.
//
// Percentages in the spec use the `ST_Percentage` style — `100000`
// represents 100% — though some third-party tools emit bare floats; we
// accept both forms.

type ColorTransformOp =
  | {
      readonly kind:
        | 'lumMod'
        | 'lumOff'
        | 'shade'
        | 'tint'
        | 'satMod'
        | 'satOff'
        | 'hueMod'
        | 'hueOff'
        | 'alpha'
        | 'alphaMod'
        | 'alphaOff';
      readonly val: number;
    }
  | { readonly kind: 'gray' | 'inv' | 'comp' };

const COLOR_TRANSFORM_LOCALS: ReadonlySet<string> = new Set([
  'lumMod',
  'lumOff',
  'shade',
  'tint',
  'satMod',
  'satOff',
  'hueMod',
  'hueOff',
  'alpha',
  'alphaMod',
  'alphaOff',
  'gray',
  'inv',
  'comp',
]);

const parseColorTransforms = (colorEl: XmlElement): readonly ColorTransformOp[] => {
  const out: ColorTransformOp[] = [];
  for (const child of colorEl.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.dml) continue;
    const local = child.name.localName;
    if (!COLOR_TRANSFORM_LOCALS.has(local)) continue;
    if (local === 'gray' || local === 'inv' || local === 'comp') {
      out.push({ kind: local });
      continue;
    }
    const raw = getAttrValue(child, qname('', 'val', ''));
    if (raw === null) continue;
    let n = Number.parseFloat(raw);
    if (!Number.isFinite(n)) continue;
    // PowerPoint emits ST_Percentage (`100000` = 100%); tolerate the
    // bare-float form some third-party tools emit.
    if (Math.abs(n) > 1) n = n / 100000;
    out.push({ kind: local as Exclude<ColorTransformOp['kind'], 'gray' | 'inv' | 'comp'>, val: n });
  }
  return out;
};

const hexToRgb01 = (hex: string): [number, number, number] => {
  const h = hex.startsWith('#') ? hex.slice(1) : hex;
  return [
    Number.parseInt(h.slice(0, 2), 16) / 255,
    Number.parseInt(h.slice(2, 4), 16) / 255,
    Number.parseInt(h.slice(4, 6), 16) / 255,
  ];
};

const rgb01ToHex = (r: number, g: number, b: number): string => {
  const clamp = (v: number): number => Math.max(0, Math.min(255, Math.round(v * 255)));
  const part = (n: number): string => n.toString(16).padStart(2, '0').toUpperCase();
  return `#${part(clamp(r))}${part(clamp(g))}${part(clamp(b))}`;
};

const rgbToHsl = (r: number, g: number, b: number): [number, number, number] => {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
};

const hueToRgb = (p: number, q: number, t: number): number => {
  if (t < 0) t += 1;
  if (t > 1) t -= 1;
  if (t < 1 / 6) return p + (q - p) * 6 * t;
  if (t < 1 / 2) return q;
  if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
  return p;
};

const hslToRgb = (h: number, s: number, l: number): [number, number, number] => {
  if (s === 0) return [l, l, l];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [hueToRgb(p, q, h + 1 / 3), hueToRgb(p, q, h), hueToRgb(p, q, h - 1 / 3)];
};

const applyColorTransforms = (hex: string, transforms: readonly ColorTransformOp[]): string => {
  if (transforms.length === 0) return hex;
  let [r, g, b] = hexToRgb01(hex);
  for (const t of transforms) {
    switch (t.kind) {
      case 'inv':
        r = 1 - r;
        g = 1 - g;
        b = 1 - b;
        break;
      case 'gray': {
        const y = 0.3 * r + 0.59 * g + 0.11 * b;
        r = g = b = y;
        break;
      }
      case 'comp': {
        const [h, s, l] = rgbToHsl(r, g, b);
        [r, g, b] = hslToRgb((h + 0.5) % 1, s, l);
        break;
      }
      case 'shade':
        // Mix toward black: out = base * val
        r *= t.val;
        g *= t.val;
        b *= t.val;
        break;
      case 'tint':
        // Mix toward white: out = base * val + (1 - val)
        r = r * t.val + (1 - t.val);
        g = g * t.val + (1 - t.val);
        b = b * t.val + (1 - t.val);
        break;
      case 'lumMod':
      case 'lumOff': {
        const [h, s, l] = rgbToHsl(r, g, b);
        const newL = Math.max(0, Math.min(1, t.kind === 'lumMod' ? l * t.val : l + t.val));
        [r, g, b] = hslToRgb(h, s, newL);
        break;
      }
      case 'satMod':
      case 'satOff': {
        const [h, s, l] = rgbToHsl(r, g, b);
        const newS = Math.max(0, Math.min(1, t.kind === 'satMod' ? s * t.val : s + t.val));
        [r, g, b] = hslToRgb(h, newS, l);
        break;
      }
      case 'hueMod':
      case 'hueOff': {
        const [h, s, l] = rgbToHsl(r, g, b);
        const newH = (((t.kind === 'hueMod' ? h * t.val : h + t.val / 360) % 1) + 1) % 1;
        [r, g, b] = hslToRgb(newH, s, l);
        break;
      }
      // alpha / alphaMod / alphaOff intentionally don't touch RGB — they
      // surface as `fill-opacity`, not as a tinted color.
    }
  }
  return rgb01ToHex(r, g, b);
};

const SCHEME_TOKEN_TO_THEME_KEY: Record<string, keyof Omit<PresentationTheme, 'name'>> = {
  tx1: 'dark1',
  dk1: 'dark1',
  bg1: 'light1',
  lt1: 'light1',
  tx2: 'dark2',
  dk2: 'dark2',
  bg2: 'light2',
  lt2: 'light2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hyperlink',
  folHlink: 'followedHyperlink',
};

const resolveSchemeToken = (token: string, theme: PresentationTheme | null): string | null => {
  if (!theme) return null;
  const key = SCHEME_TOKEN_TO_THEME_KEY[token];
  if (!key) return null;
  const hex = theme[key];
  if (typeof hex !== 'string') return null;
  const normalized = hex.startsWith('#') ? hex : `#${hex}`;
  return /^#[0-9A-Fa-f]{6}$/.test(normalized) ? normalized.toUpperCase() : null;
};

/**
 * Resolves a DrawingML color element (`<a:srgbClr>` / `<a:schemeClr>` /
 * `<a:sysClr>` / `<a:prstClr>`) with all its `<a:lumMod>` / `<a:tint>` /
 * `<a:shade>` / `<a:satMod>` etc. transform children applied. Returns
 * `null` when the color is a scheme token and no theme is supplied to
 * resolve it.
 *
 * Exposed because both run-format and fill-format code paths need to
 * apply the same transform pipeline; keeping a single implementation
 * means future spec-coverage additions only have to land in one place.
 */
export const resolveDrawingColor = (
  colorEl: XmlElement,
  theme: PresentationTheme | null,
): string | null => {
  if (colorEl.name.namespaceURI !== NS.dml) return null;
  const local = colorEl.name.localName;
  let baseHex: string | null = null;
  if (local === 'srgbClr') {
    const v = getAttrValue(colorEl, qname('', 'val', ''));
    if (v) baseHex = `#${v.toUpperCase()}`;
  } else if (local === 'schemeClr') {
    const v = getAttrValue(colorEl, qname('', 'val', ''));
    if (v) baseHex = resolveSchemeToken(v, theme);
  } else if (local === 'sysClr') {
    const last = getAttrValue(colorEl, qname('', 'lastClr', ''));
    if (last) baseHex = `#${last.toUpperCase()}`;
  } else if (local === 'prstClr') {
    // Preset colors aren't worth a full lookup table in this pass —
    // black / white cover most cases anyone reaches for in PresentationML.
    const v = getAttrValue(colorEl, qname('', 'val', ''));
    if (v === 'black') baseHex = '#000000';
    else if (v === 'white') baseHex = '#FFFFFF';
  }
  if (!baseHex) return null;
  return applyColorTransforms(baseHex, parseColorTransforms(colorEl));
};

// Reads any element shaped like `CT_TextCharacterProperties` (the schema
// shared by `<a:rPr>`, `<a:defRPr>`, and `<a:endParaRPr>`) into a partial
// TextFormat. Used by both the literal-only `getShapeRunFormat` and the
// inheritance-aware `getShapeRunFormatEffective`.
//
// When `ctx.theme` is provided, scheme tokens are resolved to concrete
// `#RRGGBB` and color transforms (`<a:lumMod>` etc.) are applied. Without
// a theme, transforms are not applied and theme tokens are passed through
// verbatim — this preserves the legacy `getShapeRunFormat` behavior.
const parseRPrLikeElement = (
  rPr: XmlElement,
  ctx?: { readonly theme: PresentationTheme | null },
): Partial<TextFormat> => {
  const out: Partial<TextFormat> = {};
  const sz = getAttrValue(rPr, qname('', 'sz', ''));
  if (sz !== null) {
    const n = Number.parseInt(sz, 10);
    if (Number.isFinite(n)) out.size = n / 100;
  }
  const b = getAttrValue(rPr, qname('', 'b', ''));
  if (b !== null) out.bold = b !== '0';
  const i = getAttrValue(rPr, qname('', 'i', ''));
  if (i !== null) out.italic = i !== '0';
  const u = getAttrValue(rPr, qname('', 'u', ''));
  if (u !== null) {
    if (u === 'none') out.underline = false;
    else if (u === 'sng') out.underline = true;
    else out.underline = u;
  }
  const strike = getAttrValue(rPr, qname('', 'strike', ''));
  if (strike !== null) {
    if (strike === 'noStrike') out.strike = false;
    else if (strike === 'sngStrike') out.strike = true;
    else out.strike = strike;
  }
  const spc = getAttrValue(rPr, qname('', 'spc', ''));
  if (spc !== null) {
    const n = Number.parseInt(spc, 10);
    if (Number.isFinite(n)) out.spc = n;
  }
  const kern = getAttrValue(rPr, qname('', 'kern', ''));
  if (kern !== null) {
    const n = Number.parseInt(kern, 10);
    if (Number.isFinite(n)) out.kern = n;
  }
  const baselineAttr = getAttrValue(rPr, qname('', 'baseline', ''));
  if (baselineAttr !== null) {
    // ST_Percentage: 100000 = 100%; tolerate bare floats.
    let n = Number.parseFloat(baselineAttr);
    if (Number.isFinite(n)) {
      if (Math.abs(n) > 1) n = n / 100000;
      out.baseline = n;
    }
  }
  const cap = getAttrValue(rPr, qname('', 'cap', ''));
  if (cap === 'none' || cap === 'small' || cap === 'all') {
    out.cap = cap;
  }
  // <a:highlight><a:srgbClr val="…"/></a:highlight>
  const highlight = firstChildElement(rPr, qname('a', 'highlight', NS.dml));
  if (highlight !== null) {
    let hlChild: XmlElement | null = null;
    for (const c of highlight.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      hlChild = c;
      break;
    }
    if (hlChild) {
      if (ctx) {
        const hex = resolveDrawingColor(hlChild, ctx.theme);
        if (hex !== null) out.highlight = hex;
      } else if (hlChild.name.localName === 'srgbClr') {
        const v = getAttrValue(hlChild, qname('', 'val', ''));
        if (v !== null) out.highlight = `#${v.toUpperCase()}`;
      } else if (hlChild.name.localName === 'schemeClr') {
        const v = getAttrValue(hlChild, qname('', 'val', ''));
        if (v !== null) out.highlight = v;
      }
    }
  }
  const solidFill = firstChildElement(rPr, qname('a', 'solidFill', NS.dml));
  if (solidFill !== null) {
    // Find the inner color element (srgbClr / schemeClr / sysClr / prstClr).
    // CT_SolidColorFillProperties holds exactly one EG_ColorChoice child.
    let colorChild: XmlElement | null = null;
    for (const c of solidFill.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      colorChild = c;
      break;
    }
    if (colorChild) {
      if (ctx) {
        // Apply transforms + resolve scheme tokens to hex.
        const hex = resolveDrawingColor(colorChild, ctx.theme);
        if (hex !== null) out.color = hex;
        else if (colorChild.name.localName === 'schemeClr') {
          // Theme not provided / token not in scheme — surface the raw token.
          const v = getAttrValue(colorChild, qname('', 'val', ''));
          if (v !== null) out.color = v;
        }
      } else {
        // Legacy `getShapeRunFormat` path: no transforms, scheme tokens
        // emitted as bare strings to match prior public behavior.
        if (colorChild.name.localName === 'srgbClr') {
          const v = getAttrValue(colorChild, qname('', 'val', ''));
          if (v !== null) out.color = `#${v.toUpperCase()}`;
        } else if (colorChild.name.localName === 'schemeClr') {
          const v = getAttrValue(colorChild, qname('', 'val', ''));
          if (v !== null) out.color = v;
        }
      }
    }
  }
  const latin = firstChildElement(rPr, qname('a', 'latin', NS.dml));
  if (latin !== null) {
    const t = getAttrValue(latin, qname('', 'typeface', ''));
    if (t !== null) out.font = t;
  }
  return out;
};

/**
 * Reads back the format of a single run. Returns `null` when the run
 * has no `<a:rPr>` (it inherits its format from the paragraph /
 * layout / master). Boolean attributes that are explicitly `"0"`
 * decode to `false`.
 *
 * Use `getShapeRunFormatEffective` if you want the resolved format
 * after walking the placeholder / lstStyle / master inheritance chain.
 */
export const getShapeRunFormat = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): TextFormat | null => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  const rPr = firstChildElement(run, NAME_A_RPR);
  if (rPr === null) return null;
  return parseRPrLikeElement(rPr) as TextFormat;
};

// -- Effective rPr cascade (ECMA-376 §21.1.2.4.7) ---------------------------
//
// A run's effective character properties are resolved by walking the
// inheritance chain — each level fills in fields that no earlier level
// supplied. First-wins per property:
//
//   1. The run's own `<a:rPr>`
//   2. The paragraph's `<a:endParaRPr>` (last run only)
//   3. The paragraph's `<a:pPr><a:defRPr>` (paragraph-level run defaults)
//   4. The text body's `<a:lstStyle><a:lvl{N+1}pPr><a:defRPr>` (N = paragraph level)
//   5. The same path on the matching placeholder in the slide's layout
//   6. The same path on the matching placeholder on the slide master,
//      then the master's `<p:txStyles>` (`titleStyle` / `bodyStyle` / `otherStyle`)
//   7. The theme's `<a:fontScheme>` — font typeface fallback only
//
// Placeholder matching: by `<p:ph/@idx>` first, then by `<p:ph/@type>`.

const NAME_A_DEF_RPR = qname('a', 'defRPr', NS.dml);
const NAME_A_END_PARA_RPR = qname('a', 'endParaRPr', NS.dml);
const NAME_A_LST_STYLE = qname('a', 'lstStyle', NS.dml);
const NAME_P_TX_BODY_PML = qname('p', 'txBody', NS.pml);
const NAME_P_TX_STYLES = qname('p', 'txStyles', NS.pml);
const NAME_P_TITLE_STYLE = qname('p', 'titleStyle', NS.pml);
const NAME_P_BODY_STYLE = qname('p', 'bodyStyle', NS.pml);
const NAME_P_OTHER_STYLE = qname('p', 'otherStyle', NS.pml);

const mergeRPrLayer = (base: Partial<TextFormat>, layer: Partial<TextFormat>): void => {
  if (base.font === undefined && layer.font !== undefined) base.font = layer.font;
  if (base.size === undefined && layer.size !== undefined) base.size = layer.size;
  if (base.color === undefined && layer.color !== undefined) base.color = layer.color;
  if (base.bold === undefined && layer.bold !== undefined) base.bold = layer.bold;
  if (base.italic === undefined && layer.italic !== undefined) base.italic = layer.italic;
  if (base.underline === undefined && layer.underline !== undefined) {
    base.underline = layer.underline;
  }
  if (base.strike === undefined && layer.strike !== undefined) base.strike = layer.strike;
  if (base.spc === undefined && layer.spc !== undefined) base.spc = layer.spc;
  if (base.kern === undefined && layer.kern !== undefined) base.kern = layer.kern;
  if (base.baseline === undefined && layer.baseline !== undefined) base.baseline = layer.baseline;
  if (base.cap === undefined && layer.cap !== undefined) base.cap = layer.cap;
  if (base.highlight === undefined && layer.highlight !== undefined) {
    base.highlight = layer.highlight;
  }
};

// `<a:lstStyle>` carries one `<a:lvl{N}pPr>` per outline level (1..9, plus
// `<a:defPPr>` for the level-0 default). Returns the inner `<a:defRPr>` for
// the requested zero-based level, or `null` if the level isn't authored.
const lstStyleLevelDefRPr = (lstStyle: XmlElement | null, level: number): XmlElement | null => {
  if (!lstStyle) return null;
  const localName = `lvl${Math.max(0, Math.min(8, level)) + 1}pPr`;
  const lvlPPr = firstChildElement(lstStyle, qname('a', localName, NS.dml));
  if (!lvlPPr) {
    // Fall back to `<a:defPPr>` only for level 0 — that's what the schema
    // declares as the "no explicit level" slot.
    if (level !== 0) return null;
    const defPPr = firstChildElement(lstStyle, qname('a', 'defPPr', NS.dml));
    if (!defPPr) return null;
    return firstChildElement(defPPr, NAME_A_DEF_RPR);
  }
  return firstChildElement(lvlPPr, NAME_A_DEF_RPR);
};

// Companion to `lstStyleLevelDefRPr` but returns the `<a:lvlNpPr>` (or
// `<a:defPPr>` for level 0) element itself — i.e. the paragraph-property
// container, not the run-default child. Used by the pPr cascade.
const lstStyleLevelPPr = (lstStyle: XmlElement | null, level: number): XmlElement | null => {
  if (!lstStyle) return null;
  const localName = `lvl${Math.max(0, Math.min(8, level)) + 1}pPr`;
  const lvlPPr = firstChildElement(lstStyle, qname('a', localName, NS.dml));
  if (lvlPPr) return lvlPPr;
  if (level !== 0) return null;
  return firstChildElement(lstStyle, qname('a', 'defPPr', NS.dml));
};

const findShapeLstStyleElement = (shape: SlideShapeData): XmlElement | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_P_TX_BODY_PML);
  if (!txBody) return null;
  return firstChildElement(txBody, NAME_A_LST_STYLE);
};

const findPlaceholderShapeIn = (
  shapes: ReadonlyArray<{
    placeholderIdx: number | null;
    placeholderType: string | null;
    element: XmlElement;
  }>,
  phIdx: number | null,
  phType: string | null,
): { element: XmlElement } | undefined => {
  let match = phIdx !== null ? shapes.find((s) => s.placeholderIdx === phIdx) : undefined;
  if (!match && phType !== null) {
    match = shapes.find((s) => s.placeholderType === phType);
  }
  return match;
};

const extractPlaceholderLstStyle = (placeholderEl: XmlElement): XmlElement | null => {
  const txBody = firstChildElement(placeholderEl, NAME_P_TX_BODY_PML);
  if (!txBody) return null;
  return firstChildElement(txBody, NAME_A_LST_STYLE);
};

const masterTxStyleFor = (masterRoot: XmlElement, phType: string | null): XmlElement | null => {
  const txStyles = firstChildElement(masterRoot, NAME_P_TX_STYLES);
  if (!txStyles) return null;
  if (phType === 'title' || phType === 'ctrTitle') {
    return firstChildElement(txStyles, NAME_P_TITLE_STYLE);
  }
  // Body / null-typed (= body default) / subTitle all inherit from bodyStyle.
  if (phType === 'body' || phType === 'subTitle' || phType === null) {
    return firstChildElement(txStyles, NAME_P_BODY_STYLE);
  }
  // Footer / date / sldNum / etc. inherit from otherStyle.
  return firstChildElement(txStyles, NAME_P_OTHER_STYLE);
};

/**
 * Resolves a run's effective character properties by walking the
 * ECMA-376 §21.1.2.4.7 inheritance chain — run rPr → endParaRPr →
 * pPr defRPr → text-body lstStyle → layout placeholder lstStyle →
 * master placeholder lstStyle + master txStyles → theme fontScheme.
 *
 * Each property (font, size, color, bold, italic, underline) is
 * resolved independently: the innermost layer that supplies a value
 * wins for that one property.
 *
 * Returns a non-null `TextFormat`; fields the cascade couldn't
 * resolve are simply absent (the renderer falls back to placeholder
 * defaults).
 *
 * Use `getShapeRunFormat` if you only want the literal `<a:rPr>` on
 * the run without inheritance.
 */
export const getShapeRunFormatEffective = (
  pres: PresentationData,
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): TextFormat => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const run = requireRun(shape, paragraphIndex, runIndex);
  const result: Partial<TextFormat> = {};

  // Theme is consulted (a) at each layer to resolve scheme tokens and
  // color transforms eagerly, so the cascade can pick the innermost layer
  // that produces a concrete color, and (b) for typeface fallback at
  // layer 7. Reading once up-front keeps the per-layer cost flat.
  const theme = getPresentationTheme(pres);
  const ctx = { theme } as const;

  // Paragraph level (0..8). `<a:pPr lvl="..">`; absent = 0.
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  let level = 0;
  if (pPr) {
    const lvlAttr = getAttrValue(pPr, ATTR_LVL);
    if (lvlAttr !== null) {
      const parsed = Number.parseInt(lvlAttr, 10);
      if (Number.isFinite(parsed)) level = parsed;
    }
  }

  // 1. Run's own rPr.
  const runRPr = firstChildElement(run, NAME_A_RPR);
  if (runRPr) mergeRPrLayer(result, parseRPrLikeElement(runRPr, ctx));

  // 2. endParaRPr — applies to the last run in the paragraph per the spec.
  const runs = runsOf(paragraph);
  if (runs.length > 0 && runs[runs.length - 1] === run) {
    const endRPr = firstChildElement(paragraph, NAME_A_END_PARA_RPR);
    if (endRPr) mergeRPrLayer(result, parseRPrLikeElement(endRPr, ctx));
  }

  // 3. Paragraph-level defaults (pPr/defRPr).
  if (pPr) {
    const defRPr = firstChildElement(pPr, NAME_A_DEF_RPR);
    if (defRPr) mergeRPrLayer(result, parseRPrLikeElement(defRPr, ctx));
  }

  // 4. Text-body lstStyle at the paragraph's level.
  const shapeLstStyle = findShapeLstStyleElement(shape);
  const shapeLvlDef = lstStyleLevelDefRPr(shapeLstStyle, level);
  if (shapeLvlDef) mergeRPrLayer(result, parseRPrLikeElement(shapeLvlDef, ctx));

  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);

  const slide = shape[SHAPE_SLIDE];
  const layout = getSlideLayout(slide);

  if (layout) {
    // 5. Matching placeholder on the layout — both its inline rPr-bearing
    //    paragraph children (if the layout authored prompt text) and its
    //    own lstStyle.
    const layoutPh = findPlaceholderShapeIn(layout[LAYOUT_PART].shapes, phIdx, phType);
    if (layoutPh) {
      const layoutLst = extractPlaceholderLstStyle(layoutPh.element);
      const layoutLvlDef = lstStyleLevelDefRPr(layoutLst, level);
      if (layoutLvlDef) mergeRPrLayer(result, parseRPrLikeElement(layoutLvlDef, ctx));
    }

    // 6. Walk one rel up to the slide master.
    const pkg = pres[INTERNAL_PACKAGE];
    const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
    const layoutRels = pkg.getRels(layoutPartName);
    if (layoutRels) {
      const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
      if (masterRel) {
        const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
        if (masterPart) {
          const masterRoot = parseXml(decode(masterPart.data)).root;
          const { shapes: masterShapes } = readShapeTreeFromCsldRoot(masterRoot, 'sldMaster');
          const masterPh = findPlaceholderShapeIn(masterShapes, phIdx, phType);
          if (masterPh) {
            const masterLst = extractPlaceholderLstStyle(masterPh.element);
            const masterLvlDef = lstStyleLevelDefRPr(masterLst, level);
            if (masterLvlDef) mergeRPrLayer(result, parseRPrLikeElement(masterLvlDef, ctx));
          }
          // Master text-style defaults (title / body / other).
          const txStyle = masterTxStyleFor(masterRoot, phType);
          const txLvlDef = lstStyleLevelDefRPr(txStyle, level);
          if (txLvlDef) mergeRPrLayer(result, parseRPrLikeElement(txLvlDef, ctx));
        }
      }
    }
  }

  // 7. Theme fontScheme — typeface resolution.
  //
  // The master often writes its `<a:latin typeface="+mj-lt"/>` /
  // `+mn-lt` placeholder tokens instead of a concrete face. Those
  // tokens must be resolved against the theme to produce a real
  // typeface; otherwise renderers see literal `+mj-lt` and fall
  // back to a generic font.
  //
  // When no layer in the cascade supplied a font at all, pick the
  // major font for title-class placeholders and the minor font for
  // everything else, matching PowerPoint's defaults.
  const fonts = getPresentationFonts(pres);
  if (fonts) {
    const resolveThemeToken = (token: string): string | undefined => {
      switch (token) {
        case '+mj-lt':
          return fonts.majorLatin ?? undefined;
        case '+mn-lt':
          return fonts.minorLatin ?? undefined;
        case '+mj-ea':
          return fonts.majorEastAsian ?? undefined;
        case '+mn-ea':
          return fonts.minorEastAsian ?? undefined;
        case '+mj-cs':
          return fonts.majorComplexScript ?? undefined;
        case '+mn-cs':
          return fonts.minorComplexScript ?? undefined;
        default:
          return undefined;
      }
    };
    if (typeof result.font === 'string' && result.font.startsWith('+')) {
      const resolved = resolveThemeToken(result.font);
      if (resolved) result.font = resolved;
    }
    if (result.font === undefined) {
      const useMajor = phType === 'title' || phType === 'ctrTitle';
      const fallback = useMajor ? fonts.majorLatin : fonts.minorLatin;
      if (fallback) result.font = fallback;
    }
  }

  return result as TextFormat;
};

// -- Effective pPr cascade --------------------------------------------------
//
// Mirror of the rPr cascade for paragraph-level properties: alignment,
// indents, line spacing, paragraph spacing, rtl. Walks the same layers:
//
//   1. The paragraph's own `<a:pPr>`
//   2. The text body's `<a:lstStyle><a:lvl{N+1}pPr>` (paragraph defaults)
//   3. The matching layout placeholder's lstStyle
//   4. The matching master placeholder's lstStyle, then
//      `<p:txStyles>/{title|body|other}Style/<a:lvl{N+1}pPr>`
//
// Each property merges independently — innermost layer that supplies a
// value wins for that one property.

/** Effective paragraph properties returned by `getParagraphPropertiesEffective`. */
export interface ParagraphProperties {
  /** Horizontal alignment per `ParagraphAlignment`. */
  align: ParagraphAlignment | null;
  /** Outline level (0..8). 0 = top-level paragraph. */
  level: number;
  /** Left indent in EMU. */
  marL: number | null;
  /** Right indent in EMU. */
  marR: number | null;
  /** First-line indent in EMU; negative for hanging indents. */
  indent: number | null;
  /** Line spacing — either a percent multiplier or a fixed point value. */
  lineSpacing:
    | { readonly kind: 'pct'; readonly value: number }
    | { readonly kind: 'pts'; readonly value: number }
    | null;
  /** Space before the paragraph in points. */
  spcBefPts: number | null;
  /** Space after the paragraph in points. */
  spcAftPts: number | null;
  /** Right-to-left paragraph (`<a:pPr rtl="1"/>`). */
  rtl: boolean | null;
}

const ALIGN_TOKEN_MAP: Record<string, ParagraphProperties['align']> = {
  l: 'left',
  ctr: 'center',
  r: 'right',
  just: 'justify',
  justLow: 'justify',
  dist: 'distribute',
  thaiDist: 'distribute',
};

const parsePPrLikeElement = (pPr: XmlElement): Partial<ParagraphProperties> => {
  const out: Partial<ParagraphProperties> = {};
  const algn = getAttrValue(pPr, qname('', 'algn', ''));
  if (algn !== null && ALIGN_TOKEN_MAP[algn] !== undefined) out.align = ALIGN_TOKEN_MAP[algn];
  const marL = getAttrValue(pPr, qname('', 'marL', ''));
  if (marL !== null) {
    const n = Number.parseInt(marL, 10);
    if (Number.isFinite(n)) out.marL = n;
  }
  const marR = getAttrValue(pPr, qname('', 'marR', ''));
  if (marR !== null) {
    const n = Number.parseInt(marR, 10);
    if (Number.isFinite(n)) out.marR = n;
  }
  const indent = getAttrValue(pPr, qname('', 'indent', ''));
  if (indent !== null) {
    const n = Number.parseInt(indent, 10);
    if (Number.isFinite(n)) out.indent = n;
  }
  const rtl = getAttrValue(pPr, qname('', 'rtl', ''));
  if (rtl !== null) out.rtl = rtl === '1' || rtl === 'true';
  const lnSpc = firstChildElement(pPr, qname('a', 'lnSpc', NS.dml));
  if (lnSpc) {
    const pct = firstChildElement(lnSpc, qname('a', 'spcPct', NS.dml));
    if (pct) {
      const v = getAttrValue(pct, qname('', 'val', ''));
      if (v !== null) {
        let n = Number.parseFloat(v);
        if (Number.isFinite(n)) {
          if (Math.abs(n) > 1) n = n / 100000;
          out.lineSpacing = { kind: 'pct', value: n };
        }
      }
    } else {
      const pts = firstChildElement(lnSpc, qname('a', 'spcPts', NS.dml));
      if (pts) {
        const v = getAttrValue(pts, qname('', 'val', ''));
        if (v !== null) {
          const n = Number.parseInt(v, 10);
          if (Number.isFinite(n)) out.lineSpacing = { kind: 'pts', value: n / 100 };
        }
      }
    }
  }
  const readSpcSide = (local: 'spcBef' | 'spcAft'): number | null => {
    const side = firstChildElement(pPr, qname('a', local, NS.dml));
    if (!side) return null;
    const pts = firstChildElement(side, qname('a', 'spcPts', NS.dml));
    if (!pts) return null;
    const v = getAttrValue(pts, qname('', 'val', ''));
    if (v === null) return null;
    const n = Number.parseInt(v, 10);
    return Number.isFinite(n) ? n / 100 : null;
  };
  const before = readSpcSide('spcBef');
  if (before !== null) out.spcBefPts = before;
  const after = readSpcSide('spcAft');
  if (after !== null) out.spcAftPts = after;
  return out;
};

const mergePPrLayer = (
  base: Partial<ParagraphProperties>,
  layer: Partial<ParagraphProperties>,
): void => {
  if (base.align === undefined && layer.align !== undefined) base.align = layer.align;
  if (base.marL === undefined && layer.marL !== undefined) base.marL = layer.marL;
  if (base.marR === undefined && layer.marR !== undefined) base.marR = layer.marR;
  if (base.indent === undefined && layer.indent !== undefined) base.indent = layer.indent;
  if (base.rtl === undefined && layer.rtl !== undefined) base.rtl = layer.rtl;
  if (base.lineSpacing === undefined && layer.lineSpacing !== undefined) {
    base.lineSpacing = layer.lineSpacing;
  }
  if (base.spcBefPts === undefined && layer.spcBefPts !== undefined) {
    base.spcBefPts = layer.spcBefPts;
  }
  if (base.spcAftPts === undefined && layer.spcAftPts !== undefined) {
    base.spcAftPts = layer.spcAftPts;
  }
};

/**
 * Resolves a paragraph's effective properties by walking the same
 * inheritance chain `getShapeRunFormatEffective` uses, but for the
 * paragraph-level surface:
 *
 *   - alignment, indent (left / right / first-line), line spacing,
 *     paragraph spacing (before / after), rtl.
 *
 * Each property is resolved independently; the innermost layer that
 * sets it wins. Fields the cascade can't resolve come through as `null`
 * so renderers know to fall back to their own defaults.
 *
 * Companion to `getParagraphAlignment` / `getParagraphLineSpacing` /
 * `getParagraphIndent` / `getParagraphSpacing`, which only surface the
 * literal `<a:pPr>` and skip the layout / master cascade.
 */
export const getParagraphPropertiesEffective = (
  pres: PresentationData,
  shape: SlideShapeData,
  paragraphIndex: number,
): ParagraphProperties => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);

  let level = 0;
  if (pPr) {
    const lvlAttr = getAttrValue(pPr, ATTR_LVL);
    if (lvlAttr !== null) {
      const parsed = Number.parseInt(lvlAttr, 10);
      if (Number.isFinite(parsed)) level = parsed;
    }
  }

  const result: Partial<ParagraphProperties> = {};

  // 1. Paragraph's own pPr.
  if (pPr) mergePPrLayer(result, parsePPrLikeElement(pPr));

  // 2. Text-body lstStyle at the paragraph's level.
  const shapeLstStyle = findShapeLstStyleElement(shape);
  const shapeLvlPPr = lstStyleLevelPPr(shapeLstStyle, level);
  if (shapeLvlPPr) mergePPrLayer(result, parsePPrLikeElement(shapeLvlPPr));

  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);
  const slide = shape[SHAPE_SLIDE];
  const layout = getSlideLayout(slide);

  if (layout) {
    // 3. Layout placeholder lstStyle.
    const layoutPh = findPlaceholderShapeIn(layout[LAYOUT_PART].shapes, phIdx, phType);
    if (layoutPh) {
      const layoutLst = extractPlaceholderLstStyle(layoutPh.element);
      const layoutLvlPPr = lstStyleLevelPPr(layoutLst, level);
      if (layoutLvlPPr) mergePPrLayer(result, parsePPrLikeElement(layoutLvlPPr));
    }

    // 4. Master placeholder lstStyle + master txStyles.
    const pkg = pres[INTERNAL_PACKAGE];
    const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
    const layoutRels = pkg.getRels(layoutPartName);
    if (layoutRels) {
      const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
      if (masterRel) {
        const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
        if (masterPart) {
          const masterRoot = parseXml(decode(masterPart.data)).root;
          const { shapes: masterShapes } = readShapeTreeFromCsldRoot(masterRoot, 'sldMaster');
          const masterPh = findPlaceholderShapeIn(masterShapes, phIdx, phType);
          if (masterPh) {
            const masterLst = extractPlaceholderLstStyle(masterPh.element);
            const masterLvlPPr = lstStyleLevelPPr(masterLst, level);
            if (masterLvlPPr) mergePPrLayer(result, parsePPrLikeElement(masterLvlPPr));
          }
          const txStyle = masterTxStyleFor(masterRoot, phType);
          const txLvlPPr = lstStyleLevelPPr(txStyle, level);
          if (txLvlPPr) mergePPrLayer(result, parsePPrLikeElement(txLvlPPr));
        }
      }
    }
  }

  return {
    align: result.align ?? null,
    level,
    marL: result.marL ?? null,
    marR: result.marR ?? null,
    indent: result.indent ?? null,
    lineSpacing: result.lineSpacing ?? null,
    spcBefPts: result.spcBefPts ?? null,
    spcAftPts: result.spcAftPts ?? null,
    rtl: result.rtl ?? null,
  };
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
 * Reads the external URL the first run in the shape's text-body links
 * to (set via `setShapeHyperlink`). Returns `null` when no run carries
 * an `<a:hlinkClick r:id=…/>` or the rId resolves to a non-hyperlink
 * target.
 */
export const getShapeHyperlink = (shape: SlideShapeData): string | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'shape') return null;
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  for (const p of txBody.children) {
    if (p.kind !== 'element' || p.name.namespaceURI !== NS.dml || p.name.localName !== 'p')
      continue;
    for (const r of p.children) {
      if (r.kind !== 'element' || r.name.namespaceURI !== NS.dml || r.name.localName !== 'r')
        continue;
      const rPr = firstChildElement(r, qname('a', 'rPr', NS.dml));
      if (!rPr) continue;
      const hlink = firstChildElement(rPr, qname('a', 'hlinkClick', NS.dml));
      if (!hlink) continue;
      const rId = getAttrValue(hlink, qname('r', 'id', NS.officeDocRels));
      if (!rId) continue;
      const slide = shape[SHAPE_SLIDE];
      const rels = slide[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]);
      if (!rels) continue;
      const rel = rels.items.find((x) => x.id === rId);
      if (rel?.type === REL_TYPES.hyperlink && rel.targetMode === 'External') {
        return rel.target;
      }
    }
  }
  return null;
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
 * Copies a shape into `targetSlide`. The source XML is cloned and
 * appended to the target's `<p:spTree>`. Image rels on the source
 * shape are followed: the linked media part is referenced from the
 * target slide via a freshly allocated rId (no media bytes are
 * copied — both slides share the underlying part).
 *
 * v1 requires source and target to live in the same package
 * (`sourceShape`'s slide and `targetSlide` must share the same
 * `OpcPackage`). Cross-package copy is `importSlide` territory.
 *
 * Returns the new `SlideShapeData` on `targetSlide`.
 */
export const copyShape = (targetSlide: SlideData, sourceShape: SlideShapeData): SlideShapeData => {
  const sourceSlide = sourceShape[SHAPE_SLIDE];
  if (sourceSlide[INTERNAL_PACKAGE] !== targetSlide[INTERNAL_PACKAGE]) {
    throw new Error(
      'copyShape: source and target must be in the same package. Use importSlide for cross-deck copies.',
    );
  }
  const pkg = targetSlide[INTERNAL_PACKAGE];
  const sourceEl = sourceShape[SHAPE_ELEMENT];

  // Deep-clone the XML by serializing + re-parsing one element.
  // We wrap in a temporary parent so we can extract the cloned element
  // back out without ambient namespaces leaking from the slide root.
  const cloned = cloneXmlElement(sourceEl);

  // Allocate a fresh shape id on the target slide and overwrite the
  // cNvPr/cNvPr id attribute.
  const newId = nextShapeId(targetSlide);
  rewriteCNvPrId(cloned, newId);

  // Walk the cloned element for r:embed / r:link references. For each
  // referenced rId in the source slide's rels, copy the rel onto the
  // target slide's rels (allocating a fresh rId) and update the cloned
  // attribute. This covers picture blips + media references.
  const sourceRels = pkg.getRels(sourceSlide[SLIDE_PART_NAME]);
  if (sourceRels) {
    const targetRels = pkg.getRels(targetSlide[SLIDE_PART_NAME]) ?? emptyRels();
    const usedIds = new Set(targetRels.items.map((r) => r.id));
    rewriteRIdReferences(cloned, (oldRId) => {
      const sourceRel = sourceRels.items.find((r) => r.id === oldRId);
      if (!sourceRel) return oldRId;
      // Look for an existing rel on target with the same type+target;
      // reuse if found to avoid duplicates.
      const existing = targetRels.items.find(
        (r) =>
          r.type === sourceRel.type &&
          r.target === sourceRel.target &&
          r.targetMode === sourceRel.targetMode,
      );
      if (existing) return existing.id;
      const newRId = nextRelId([...usedIds]);
      usedIds.add(newRId);
      targetRels.items.push({ ...sourceRel, id: newRId });
      return newRId;
    });
    pkg.setRels(targetSlide[SLIDE_PART_NAME], targetRels);
  }

  return appendAndReturnNewShape(targetSlide, cloned);
};

/** Recursively clone an XML element (no parent, deep). */
const cloneXmlElement = (el: XmlElement): XmlElement => ({
  kind: 'element',
  name: el.name,
  attrs: el.attrs.map((a) => ({ name: a.name, value: a.value })),
  prefixDecls: new Map(el.prefixDecls),
  children: el.children.map((c) => {
    if (c.kind === 'element') return cloneXmlElement(c);
    return { ...c };
  }),
});

const rewriteCNvPrId = (root: XmlElement, newId: number): void => {
  const walk = (el: XmlElement): boolean => {
    if (
      el.name.namespaceURI === NS.pml &&
      el.name.localName === 'cNvPr' &&
      el.attrs.some((a) => a.name.namespaceURI === '' && a.name.localName === 'id')
    ) {
      el.attrs = el.attrs.map((a) =>
        a.name.namespaceURI === '' && a.name.localName === 'id'
          ? { name: a.name, value: String(newId) }
          : a,
      );
      return true;
    }
    for (const c of el.children) {
      if (c.kind === 'element' && walk(c)) return true;
    }
    return false;
  };
  walk(root);
};

const rewriteRIdReferences = (root: XmlElement, map: (oldRId: string) => string): void => {
  const walk = (el: XmlElement): void => {
    el.attrs = el.attrs.map((a) => {
      if (
        a.name.namespaceURI === NS.officeDocRels &&
        (a.name.localName === 'id' || a.name.localName === 'embed' || a.name.localName === 'link')
      ) {
        return { name: a.name, value: map(a.value) };
      }
      return a;
    });
    for (const c of el.children) {
      if (c.kind === 'element') walk(c);
    }
  };
  walk(root);
};

// ---------------------------------------------------------------------------
// Z-order — move shapes forward / backward inside the slide's spTree.
//
// OOXML shape z-order is just the document order of children of
// `<p:spTree>`: the first child renders behind, the last in front.
// PowerPoint's "Bring to Front" / "Send to Back" affordances translate
// directly to reordering those children.
//
// Each function targets only "real" shape children — `<p:sp>`, `<p:pic>`,
// `<p:cxnSp>`, `<p:graphicFrame>`, `<p:grpSp>`. The required
// `<p:nvGrpSpPr>` / `<p:grpSpPr>` preface stays at the top.

const SHAPE_CHILD_LOCALS = new Set(['sp', 'pic', 'cxnSp', 'graphicFrame', 'grpSp']);

const isShapeChild = (node: {
  kind: string;
  name?: { namespaceURI: string; localName: string };
}): boolean =>
  node.kind === 'element' &&
  node.name?.namespaceURI === NS.pml &&
  SHAPE_CHILD_LOCALS.has(node.name.localName);

/** Move `shape` to the end of its spTree (render in front of all others). */
export const bringShapeToFront = (shape: SlideShapeData): void => {
  const slide = shape[SHAPE_SLIDE];
  const spTree = requireSpTree(slide);
  const target = shape[SHAPE_ELEMENT];
  const idx = spTree.children.indexOf(target);
  if (idx < 0) return;
  if (idx === spTree.children.length - 1) return; // already at front
  spTree.children.splice(idx, 1);
  spTree.children.push(target);
  commitSlideData(slide);
  rebuildShapesFromDocument(slide);
};

/**
 * Move `shape` behind every other shape on the slide. The
 * `<p:nvGrpSpPr>` / `<p:grpSpPr>` preface — required by the schema —
 * stays at the top.
 */
export const sendShapeToBack = (shape: SlideShapeData): void => {
  const slide = shape[SHAPE_SLIDE];
  const spTree = requireSpTree(slide);
  const target = shape[SHAPE_ELEMENT];
  const idx = spTree.children.indexOf(target);
  if (idx < 0) return;

  // First "shape child" position — after nvGrpSpPr / grpSpPr.
  let firstShapeAt = spTree.children.length;
  for (let i = 0; i < spTree.children.length; i++) {
    const c = spTree.children[i];
    if (c && isShapeChild(c)) {
      firstShapeAt = i;
      break;
    }
  }
  if (idx <= firstShapeAt) return;
  spTree.children.splice(idx, 1);
  spTree.children.splice(firstShapeAt, 0, target);
  commitSlideData(slide);
  rebuildShapesFromDocument(slide);
};

/** Swap `shape` with the next shape sibling (move one step forward). */
export const bringShapeForward = (shape: SlideShapeData): void => {
  const slide = shape[SHAPE_SLIDE];
  const spTree = requireSpTree(slide);
  const target = shape[SHAPE_ELEMENT];
  const idx = spTree.children.indexOf(target);
  if (idx < 0) return;
  // Find next shape sibling.
  for (let i = idx + 1; i < spTree.children.length; i++) {
    const c = spTree.children[i];
    if (c && isShapeChild(c)) {
      const next = c;
      spTree.children[idx] = next;
      spTree.children[i] = target;
      commitSlideData(slide);
      rebuildShapesFromDocument(slide);
      return;
    }
  }
};

/**
 * Returns the shape's z-index among the slide's "real" shape children
 * (`<p:sp>` / `<p:pic>` / `<p:cxnSp>` / `<p:graphicFrame>` / `<p:grpSp>`),
 * skipping the required `<p:nvGrpSpPr>` / `<p:grpSpPr>` preface.
 * Higher numbers render in front.
 */
export const getShapeZIndex = (shape: SlideShapeData): number => {
  const slide = shape[SHAPE_SLIDE];
  const spTree = requireSpTree(slide);
  let i = 0;
  for (const c of spTree.children) {
    if (!isShapeChild(c)) continue;
    if (c === shape[SHAPE_ELEMENT]) return i;
    i++;
  }
  return -1;
};

/**
 * Moves the shape to a specific z-index among the slide's "real"
 * shape children. Index is clamped to the available range. Higher
 * numbers render in front. The required preface elements stay at the
 * top of `<p:spTree>`.
 */
export const setShapeZIndex = (shape: SlideShapeData, toIndex: number): void => {
  const slide = shape[SHAPE_SLIDE];
  const spTree = requireSpTree(slide);
  const target = shape[SHAPE_ELEMENT];
  const allShapeChildren = spTree.children.filter((c): c is XmlElement => isShapeChild(c));
  const clamped = Math.max(0, Math.min(toIndex, allShapeChildren.length - 1));

  // Remove the target from the tree, then re-insert at the position
  // corresponding to z-index `clamped` among the remaining shapes.
  spTree.children = spTree.children.filter((c) => c !== target);
  const remainingShapes = spTree.children.filter((c): c is XmlElement => isShapeChild(c));
  if (clamped >= remainingShapes.length) {
    spTree.children.push(target);
  } else {
    const anchor = remainingShapes[clamped]!;
    const anchorIdx = spTree.children.indexOf(anchor);
    spTree.children.splice(anchorIdx, 0, target);
  }
  commitSlideData(slide);
  rebuildShapesFromDocument(slide);
};

/** Swap `shape` with the previous shape sibling (move one step backward). */
export const sendShapeBackward = (shape: SlideShapeData): void => {
  const slide = shape[SHAPE_SLIDE];
  const spTree = requireSpTree(slide);
  const target = shape[SHAPE_ELEMENT];
  const idx = spTree.children.indexOf(target);
  if (idx < 0) return;
  for (let i = idx - 1; i >= 0; i--) {
    const c = spTree.children[i];
    if (c && isShapeChild(c)) {
      const prev = c;
      spTree.children[idx] = prev;
      spTree.children[i] = target;
      commitSlideData(slide);
      rebuildShapesFromDocument(slide);
      return;
    }
  }
};

/**
 * Removes the shape from its slide's shape tree. Subsequent property
 * reads on this handle reflect the stale snapshot — discard it after.
 *
 * Removing a picture does NOT delete the underlying media part — it
 * may be referenced from other slides.
 */
/**
 * Removes every shape (sp / pic / cxnSp / graphicFrame / grpSp) from
 * the slide's `<p:spTree>`. The required `<p:nvGrpSpPr>` and
 * `<p:grpSpPr>` preface stays in place, so the slide is still valid
 * and re-applies its layout's placeholders on the next open.
 *
 * Useful for "start this slide over but keep its layout binding."
 */
export const clearSlideShapes = (slide: SlideData): void => {
  const spTree = requireSpTree(slide);
  spTree.children = spTree.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.pml &&
        SHAPE_CHILD_LOCALS.has(c.name.localName)
      ),
  );
  commitSlideData(slide);
  rebuildShapesFromDocument(slide);
};

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
    throw new Error(
      'addSlideImage: could not detect image format. Pass options.format explicitly.',
    );
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

  const hasDefault = pkg.contentTypes.defaults.some((d) => d.extension.toLowerCase() === extension);
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
