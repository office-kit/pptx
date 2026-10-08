// Slide layouts.

import {
  type PlaceholderType,
  type SlideLayoutType,
  readSlideLayoutPart,
} from '../../internal/presentationml/index.ts';
import { parseXml } from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  LAYOUT_DOCUMENT,
  LAYOUT_PART,
  LAYOUT_PART_NAME,
  type PresentationData,
  type SlideLayoutData,
} from '../_internal-symbols.ts';
import { SLIDE_LAYOUT_CONTENT_TYPE, decode } from './_helpers.ts';
import { type SlideLayoutPlaceholder, placeholderViews } from './_part-placeholders.ts';
import { getSlideLayout } from './shape-slide-read.ts';
import { getSlides } from './slide-query.ts';

// ---------------------------------------------------------------------------
// Slide layouts.

/** The reference desktop app's user-visible layout name. */
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

export type { SlideLayoutPlaceholder } from './_part-placeholders.ts';

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
): ReadonlyArray<SlideLayoutPlaceholder> => placeholderViews(layout[LAYOUT_PART].shapes);

/**
 * Finds the first slide layout whose user-visible name matches `name`,
 * or `null` if none does. Accepts a literal string (exact equality)
 * or a `RegExp` for pattern matches — useful when template providers
 * suffix versions onto names (`'Title and Content v2'`).
 *
 * String matching is **case-sensitive** and exact: `findSlideLayout(pres,
 * 'Blank')` matches a layout named `"Blank"` but not `"blank"`. The
 * user-visible name is also locale-dependent (a deck authored in a
 * The reference desktop app localizes `"Blank"`), so prefer
 * {@link findSlideLayoutByType} — which matches the locale-stable
 * `<p:sldLayout type="…">` token (`'blank'`, `'title'`, `'obj'`, …) —
 * when you need a robust lookup. Pass a `RegExp` with the `i` flag here
 * for a case-insensitive name match.
 */
export const findSlideLayout = (
  pres: PresentationData,
  name: string | RegExp,
): SlideLayoutData | null => {
  for (const layout of getSlideLayouts(pres)) {
    const n = layout[LAYOUT_PART].name;
    const hit = typeof name === 'string' ? n === name : name.test(n);
    if (hit) return layout;
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
  type: PlaceholderType,
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
 * across the reference desktop app's UI languages.
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
    const doc = parseXml(decode(part.data));
    out.push({
      [INTERNAL_PACKAGE]: pkg,
      [LAYOUT_PART_NAME]: part.name,
      [LAYOUT_DOCUMENT]: doc,
      [LAYOUT_PART]: readSlideLayoutPart(doc.root),
    });
  }
  return out;
};

/**
 * Layouts in the package that no slide references. Useful when
 * trimming a template deck — these layouts contribute parts and rels
 * without ever rendering. Iteration order matches `getSlideLayouts`.
 */
export const getUnusedSlideLayouts = (pres: PresentationData): ReadonlyArray<SlideLayoutData> => {
  const referenced = new Set<string>();
  for (const slide of getSlides(pres)) {
    const layout = getSlideLayout(slide);
    if (layout !== null) referenced.add(getSlideLayoutName(layout));
  }
  return getSlideLayouts(pres).filter((l) => !referenced.has(getSlideLayoutName(l)));
};

/**
 * Returns a map of layout name → number of slides that reference it.
 * Every layout enumerated by `getSlideLayouts` appears as a key (zero
 * count for unreferenced layouts), so this surfaces unused layouts
 * directly — e.g. for templates that ship with placeholder layouts the
 * deck never picks up.
 */
/**
 * Histogram of slide-layout type token (e.g. `title`, `obj`, `twoObj`,
 * `blank`) → number of slides that use a layout of that type. Useful
 * for "how many content slides vs. dividers vs. title slides?" audits
 * — keyed on the OOXML enum, so stable across locales (unlike
 * `getSlideLayoutUsageCounts`, which keys on the user-visible name).
 * Layouts with no `type` token are skipped (the spec default is
 * `cust`).
 */
export const getSlideLayoutUsageCountsByType = (
  pres: PresentationData,
): Readonly<Record<string, number>> => {
  const counts: Record<string, number> = {};
  for (const slide of getSlides(pres)) {
    const layout = getSlideLayout(slide);
    if (layout === null) continue;
    const t = layout[LAYOUT_PART].layoutType;
    if (t === null) continue;
    counts[t] = (counts[t] ?? 0) + 1;
  }
  return counts;
};

export const getSlideLayoutUsageCounts = (
  pres: PresentationData,
): Readonly<Record<string, number>> => {
  const counts: Record<string, number> = {};
  for (const layout of getSlideLayouts(pres)) {
    counts[getSlideLayoutName(layout)] = 0;
  }
  for (const slide of getSlides(pres)) {
    const layout = getSlideLayout(slide);
    if (layout === null) continue;
    const name = getSlideLayoutName(layout);
    counts[name] = (counts[name] ?? 0) + 1;
  }
  return counts;
};
