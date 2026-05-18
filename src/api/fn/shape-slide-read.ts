// Slide-level reads (shape queries on a slide).

import {
  getShapeBounds,
  getShapePreset,
  isShapePlaceholder,
  type ShapeBounds,
} from './shape-read-base.ts';
import { setShapePosition } from './shape-fill-stroke.ts';
import { replaceTokensInTree } from '../../internal/drawingml/index.ts';
import type { Emu } from '../units.ts';
import {
  basename,
  emptyRels,
  nextRelId,
  partName,
  resolveTarget,
} from '../../internal/opc/index.ts';
import {
  REL_TYPES,
  type ShapeKind,
  readSlideLayoutPart,
} from '../../internal/presentationml/index.ts';
import { parseXml, serializeXml } from '../../internal/xml/index.ts';
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
import { commitSlideData, decode, refreshSlideData } from './_helpers.ts';
import { getShapeHyperlink } from './shape-paragraph.ts';
import { getSlides } from './slide-query.ts';
import { hasShapeText } from './embedded.ts';

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
 * First shape on the slide whose `cNvPr@name` matches `name`. Accepts
 * either a literal string (exact-equality) or a `RegExp` for pattern
 * matches — mirroring `findShapeByText`. Returns `null` when nothing
 * matches.
 */
export const findShapeByName = (slide: SlideData, name: string | RegExp): SlideShapeData | null => {
  if (typeof name === 'string') {
    for (const shape of slide[SLIDE_SHAPES]) {
      if (shape[SHAPE_SNAPSHOT].name === name) return shape;
    }
    return null;
  }
  for (const shape of slide[SLIDE_SHAPES]) {
    if (name.test(shape[SHAPE_SNAPSHOT].name)) return shape;
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

/**
 * Every shape on the slide whose `cNvPr@name` matches `name`. Accepts
 * either a literal string (exact-equality, case-sensitive) or a
 * `RegExp` for pattern matches — useful when template-cloned shapes
 * share a prefix (`'TextPlaceholder1'`, `'TextPlaceholder2'`, …).
 */
export const findShapesByName = (
  slide: SlideData,
  name: string | RegExp,
): ReadonlyArray<SlideShapeData> => {
  if (typeof name === 'string') {
    return slide[SLIDE_SHAPES].filter((s) => s[SHAPE_SNAPSHOT].name === name);
  }
  return slide[SLIDE_SHAPES].filter((s) => name.test(s[SHAPE_SNAPSHOT].name));
};

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
 * Every shape on the slide whose hyperlink target matches `url`
 * (substring or `RegExp`). Pairs the existing presentation-level
 * `findSlidesByHyperlink` for cases where the caller already has a
 * specific slide and wants the linking shapes inside it.
 */
export const findShapesByHyperlink = (
  slide: SlideData,
  url: string | RegExp,
): ReadonlyArray<SlideShapeData> => {
  const out: SlideShapeData[] = [];
  for (const shape of slide[SLIDE_SHAPES]) {
    const target = getShapeHyperlink(shape);
    if (target === null) continue;
    if (typeof url === 'string' ? target.includes(url) : url.test(target)) {
      out.push(shape);
    }
  }
  return out;
};

/**
 * Every shape on the slide that carries any hyperlink (regardless of
 * target). Useful for "audit every clickable shape on this slide" —
 * counterpart to `findShapesByHyperlink(slide, url)`, which requires a
 * matching URL.
 */
export const findShapesWithHyperlinks = (slide: SlideData): ReadonlyArray<SlideShapeData> => {
  const out: SlideShapeData[] = [];
  for (const shape of slide[SLIDE_SHAPES]) {
    if (getShapeHyperlink(shape) !== null) out.push(shape);
  }
  return out;
};

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
