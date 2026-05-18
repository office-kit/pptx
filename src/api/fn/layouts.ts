// Slide layouts.

import { readPosition, readSize } from '../../internal/drawingml/index.ts';
import type { Emu } from '../units.ts';
import { type SlideLayoutType, readSlideLayoutPart } from '../../internal/presentationml/index.ts';
import { parseXml } from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  LAYOUT_PART,
  LAYOUT_PART_NAME,
  type PresentationData,
  type SlideLayoutData,
} from '../_internal-symbols.ts';
import { SLIDE_LAYOUT_CONTENT_TYPE, decode } from './_helpers.ts';
import type { ShapeBounds } from './shapes.ts';

// ---------------------------------------------------------------------------
// Slide layouts.

/** PowerPoint's user-visible layout name. */
export const getSlideLayoutName = (layout: SlideLayoutData): string => layout[LAYOUT_PART].name;

/**
 * Returns the package part name (e.g. `/ppt/slideLayouts/slideLayout3.xml`)
 * of `layout`. Useful for surfacing layouts in validator output and
 * other path-keyed UIs.
 */
export const getSlideLayoutPartName = (layout: SlideLayoutData): string => layout[LAYOUT_PART_NAME];

/**
 * Returns the slide layout whose package part name equals
 * `partName`, or `null` when no such layout exists. Mirror of
 * `findSlideByPartName` for layouts.
 */
export const findSlideLayoutByPartName = (
  pres: PresentationData,
  partName: string,
): SlideLayoutData | null => {
  for (const layout of getSlideLayouts(pres)) {
    if (layout[LAYOUT_PART_NAME] === partName) return layout;
  }
  return null;
};

/**
 * Read-only view of one placeholder on a slide layout. Surfaces the
 * three fields a slide-author cares about when binding a slide to a
 * layout: which slot is for the title, which is for the body, etc.
 */
export interface SlideLayoutPlaceholder {
  /** `<p:ph type="...">`. `null` when omitted — spec default is `body`. */
  readonly type: string | null;
  /** `<p:ph idx="...">`. `null` when omitted — spec default is `0`. */
  readonly idx: number | null;
  /** `<p:cNvPr name="...">` — what PowerPoint shows in the selection pane. */
  readonly name: string;
  /**
   * Layout-defined position + size in EMU. A slide placeholder with no
   * `<a:xfrm>` of its own inherits these. `null` when the layout
   * placeholder also lacks an explicit transform (rare — usually the
   * master defines it then).
   */
  readonly bounds: ShapeBounds | null;
}

/**
 * Enumerates the placeholder shapes on a slide layout. Non-placeholder
 * shapes (decorative rectangles, watermarks added to the layout) are
 * filtered out; only entries with a `<p:ph>` element are returned.
 *
 * Use this when you need to discover which placeholder indices a
 * layout exposes — e.g. before `findSlidePlaceholder(slide, ...)` to
 * confirm the slot exists.
 */
export const getSlideLayoutPlaceholders = (
  layout: SlideLayoutData,
): ReadonlyArray<SlideLayoutPlaceholder> => {
  const out: SlideLayoutPlaceholder[] = [];
  for (const shape of layout[LAYOUT_PART].shapes) {
    // Only `p:sp` shapes carry placeholders in real templates; pictures
    // and connectors can technically have `<p:ph>` per the schema but
    // PowerPoint never authors that. Filter for safety either way.
    if (shape.placeholderType === null && shape.placeholderIdx === null) continue;
    const pos = readPosition(shape.element, shape.kind);
    const size = readSize(shape.element, shape.kind);
    const bounds: ShapeBounds | null =
      pos === null || size === null
        ? null
        : { x: pos.x as Emu, y: pos.y as Emu, w: size.w as Emu, h: size.h as Emu };
    out.push({
      type: shape.placeholderType,
      idx: shape.placeholderIdx,
      name: shape.name,
      bounds,
    });
  }
  return out;
};

/**
 * Finds the first slide layout whose user-visible name matches `name`,
 * or `null` if none does. Convenience over `getSlideLayouts(...).find(...)`.
 */
export const findSlideLayout = (pres: PresentationData, name: string): SlideLayoutData | null => {
  for (const layout of getSlideLayouts(pres)) {
    if (layout[LAYOUT_PART].name === name) return layout;
  }
  return null;
};

/**
 * Returns every slide layout in the package that exposes a
 * placeholder of the given type token (`'title'`, `'body'`,
 * `'ftr'`, etc.). Useful for "find every layout that can host a
 * body" lookups before `addSlide`.
 */
export const findLayoutsWithPlaceholderType = (
  pres: PresentationData,
  type: string,
): ReadonlyArray<SlideLayoutData> => {
  const out: SlideLayoutData[] = [];
  for (const layout of getSlideLayouts(pres)) {
    const phs = getSlideLayoutPlaceholders(layout);
    const hit = phs.some(
      (p) => p.type === type || (type === 'body' && p.type === null && p.idx !== null),
    );
    if (hit) out.push(layout);
  }
  return out;
};

/**
 * Finds the first slide layout with the given `<p:sldLayout type="...">`
 * token. Unlike `findSlideLayout` (which matches the user-visible
 * name, and is therefore locale-sensitive), this matches the spec
 * token — `title`, `obj`, `twoObj`, `blank`, etc. — and is stable
 * across PowerPoint UI languages.
 */
export const findSlideLayoutByType = (
  pres: PresentationData,
  layoutType: SlideLayoutType | string,
): SlideLayoutData | null => {
  for (const layout of getSlideLayouts(pres)) {
    if (layout[LAYOUT_PART].layoutType === layoutType) return layout;
  }
  return null;
};

/**
 * Layout type token, when present (`title`, `obj`, `twoObj`, ...).
 * `null` when omitted — the spec default for that case is `cust`.
 */
export const getSlideLayoutType = (layout: SlideLayoutData): SlideLayoutType | string | null =>
  layout[LAYOUT_PART].layoutType;

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
