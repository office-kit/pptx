// Shape reads: identity, geometry, placeholders, bounds, group.

import { setShapePosition } from './shape-fill-stroke.ts';
import { getSlideLayout } from './shape-slide-read.ts';
import {
  type Position,
  type Size,
  readFlip,
  readPosition,
  readRotation,
  readSize,
  setPosition as writePosition,
  setSize as writeSize,
} from '../../internal/drawingml/index.ts';
import type { Emu } from '../units.ts';
import { partName, resolveTarget } from '../../internal/opc/index.ts';
import {
  REL_TYPES,
  type ShapeKind,
  readGroupChildren,
  readPresentationPart,
  readShapeTreeFromCsldRoot,
} from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
  attr,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
} from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  LAYOUT_PART,
  LAYOUT_PART_NAME,
  type PresentationData,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  SLIDE_SHAPES,
  type SlideData,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import { PRES_PART_NAME, commitAndRefresh, decode } from './_helpers.ts';
import { getSlides } from './slide-query.ts';
import { findCNvPr } from './embedded.ts';

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
