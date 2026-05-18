// Slide-level queries: enumeration, factories, search, outline, info,
// visibility, and text replacement.

import { replaceTextInTree, replaceTokensInTree } from '../../internal/drawingml/index.ts';
import { type PartName, partName } from '../../internal/opc/index.ts';
import type { OpcPackage } from '../../internal/parts/index.ts';
import {
  readPresentationPart,
  readSlidePart,
  slideText,
} from '../../internal/presentationml/index.ts';
import { NS, attr, getAttrValue, parseXml, qname, serializeXml } from '../../internal/xml/index.ts';
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
} from '../_internal-symbols.ts';
import {
  PRES_PART_NAME,
  SLIDE_CONTENT_TYPE,
  SLIDE_LAYOUT_CONTENT_TYPE,
  commitSlideData,
  decode,
  encode,
  refreshSlideData,
} from './_helpers.ts';
import { findSlidePlaceholder, getSlideLayout, getSlideShapes } from './shapes.ts';
import { getSlideNotes, setSlideNotes } from './features.ts';
import { getSlideTitle } from './embedded.ts';

// Build a SlideData handle from a part's bytes. Used by `getSlides` on
// cold load; deck mutators that change the shape count rebuild via
// `_helpers.rebuildShapesFromDocument` instead.
const buildSlideData = (pkg: OpcPackage, partNameValue: PartName, bytes: Uint8Array): SlideData => {
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
 * Enumerates slides in presentation order. Returns opaque `SlideData`
 * handles that pass to every slide-level fn-API helper.
 *
 * Throws if any referenced slide part is missing — a structurally
 * invalid PPTX cannot honor the L1 contract.
 */
/**
 * Returns the shape at the given 0-based `index` on the slide, or
 * `null` when `index` is out of range. Convenience over
 * `getSlideShapes(slide)[index] ?? null`.
 */
export const getShapeAt = (slide: SlideData, index: number): SlideShapeData | null =>
  slide[SLIDE_SHAPES][index] ?? null;

/**
 * Returns the slide count without forcing every slide part to be
 * parsed. Reads only `presentation.xml` and walks `<p:sldIdLst>`.
 *
 * Equivalent to `getSlides(pres).length`, but cheaper on cold reads
 * — useful when the caller only wants a count for UI badges or
 * validation logic.
 */
export const getSlideCount = (pres: PresentationData): number => {
  const cached = pres._slidesCache;
  if (cached !== null) return cached.length;
  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) return 0;
  const root = parseXml(decode(presPart.data)).root;
  const model = readPresentationPart(root);
  return model.slides.length;
};

/**
 * Returns the number of slide layouts in the package. Walks the
 * content-type map; cheaper than `getSlideLayouts(pres).length`
 * because it avoids parsing each layout's XML body.
 */
export const getSlideLayoutCount = (pres: PresentationData): number => {
  const pkg = pres[INTERNAL_PACKAGE];
  let n = 0;
  for (const part of pkg.parts) {
    if (part.contentType === SLIDE_LAYOUT_CONTENT_TYPE) n++;
  }
  return n;
};

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
 * Length (in code points) of the slide's concatenated visible text.
 * Counts Unicode code points via `Array.from`, so surrogate-pair
 * characters (emoji, supplementary CJK) count as 1, matching the
 * library's invariant on text editing.
 */
export const getSlideTextLength = (slide: SlideData): number =>
  Array.from(getSlideText(slide)).length;

/**
 * A single slide's outline entry: its index, title (or `null`), and
 * the text of its body placeholder (or `null`). Used by
 * `getSlideOutline`.
 */
export interface SlideOutlineEntry {
  readonly index: number;
  readonly title: string | null;
  readonly body: string | null;
}

/**
 * Returns a one-entry-per-slide outline: title + body-placeholder
 * text. Useful for building thumbnail panes, table-of-contents
 * inserts, or quick-look previews without rendering the slide.
 *
 * "Body" here is the first matching `body` placeholder per
 * `findSlidePlaceholder(slide, 'body')`. Non-placeholder text on
 * the slide isn't surfaced; use `getSlideText(slide)` for that.
 */
export const getSlideOutline = (pres: PresentationData): ReadonlyArray<SlideOutlineEntry> => {
  const out: SlideOutlineEntry[] = [];
  const slides = getSlides(pres);
  for (let i = 0; i < slides.length; i++) {
    const slide = slides[i]!;
    const body = findSlidePlaceholder(slide, 'body');
    out.push({
      index: i,
      title: getSlideTitle(slide),
      body: body !== null ? body[SHAPE_SNAPSHOT].text : null,
    });
  }
  return out;
};

/**
 * Concatenated title + body text from every slide, joined with the
 * given `separator` (defaults to `'\n\n'`). Useful for generating
 * a table-of-contents handout from a deck. Slides without a title
 * still contribute their body (and vice versa); slides without
 * both are skipped.
 */
export const getOutlineText = (pres: PresentationData, separator: string = '\n\n'): string => {
  const parts: string[] = [];
  for (const entry of getSlideOutline(pres)) {
    const segments: string[] = [];
    if (entry.title !== null) segments.push(entry.title);
    if (entry.body !== null) segments.push(entry.body);
    if (segments.length > 0) parts.push(segments.join('\n'));
  }
  return parts.join(separator);
};

/**
 * Concatenated visible text from every slide, joined with the
 * given `separator` (defaults to a form-feed, `\f`, between slides).
 * Useful for search-indexing a whole deck without iterating slides
 * yourself.
 */
export const getPresentationText = (pres: PresentationData, separator: string = '\f'): string => {
  const parts: string[] = [];
  for (const slide of getSlides(pres)) parts.push(slideText(slide[SLIDE_PART]));
  return parts.join(separator);
};

/**
 * Total code-point length of visible text across every slide.
 * Sibling of `getSlideTextLength`; counts surrogate-pair characters
 * (emoji, supplementary CJK) as 1 each. Cheaper than building the
 * concatenated string when only the length matters.
 */
export const getPresentationTextLength = (pres: PresentationData): number => {
  let n = 0;
  for (const slide of getSlides(pres)) {
    n += Array.from(slideText(slide[SLIDE_PART])).length;
  }
  return n;
};

/**
 * Dense per-slide shape count array, 0-based by slide index. Counts
 * top-level + group-children shapes (whatever `getSlideShapes` flattens).
 * Useful for charting shape density per slide and identifying outliers
 * (the 200-shape "soup" slide everybody complains about) for cleanup.
 */
export const getPresentationShapeCountsBySlide = (pres: PresentationData): ReadonlyArray<number> =>
  getSlides(pres).map((s) => s[SLIDE_SHAPES].length);

/**
 * Dense per-slide visible-text length array, 0-based by slide index.
 * Counts code points (surrogate pairs as 1) per `getSlideTextLength`.
 * Pair with `getPresentationShapeCountsBySlide` for slide-density
 * audits — high text length on a slide with few shapes usually means
 * one paragraph-heavy text box; low text on a many-shape slide
 * usually means a busy diagram.
 */
export const getPresentationTextLengthsBySlide = (pres: PresentationData): ReadonlyArray<number> =>
  getSlides(pres).map((s) => Array.from(slideText(s[SLIDE_PART])).length);

/**
 * Returns the 0-based index of `slide` within `pres`, or `-1` if the
 * slide doesn't belong to this presentation (e.g. after a removeSlide
 * call, or if it was constructed from a different package).
 */
export const getSlideIndex = (pres: PresentationData, slide: SlideData): number => {
  const slides = getSlides(pres);
  return slides.indexOf(slide);
};

/**
 * Returns the slide at the given 0-based `index`, or `null` if `index`
 * is out of range. Convenience over `getSlides(pres)[index] ?? null`.
 */
export const getSlideAt = (pres: PresentationData, index: number): SlideData | null => {
  const slides = getSlides(pres);
  return slides[index] ?? null;
};

/**
 * Per-slide summary suitable for slide pickers or audit reports.
 *
 * `index` is the 0-based position in the deck. `title` falls back to
 * `null` when the slide has no title placeholder. `layoutName` falls
 * back to `null` when the slide has no `slideLayout` rel or the
 * layout doesn't carry a user-visible name.
 */
export interface SlideInfo {
  readonly index: number;
  readonly title: string | null;
  readonly hidden: boolean;
  readonly shapeCount: number;
  readonly layoutName: string | null;
}

export const getSlideInfo = (pres: PresentationData, slide: SlideData): SlideInfo => {
  const layout = getSlideLayout(slide);
  return {
    index: getSlideIndex(pres, slide),
    title: getSlideTitle(slide),
    hidden: isSlideHidden(slide),
    shapeCount: getSlideShapes(slide).length,
    layoutName: layout ? layout[LAYOUT_PART].name : null,
  };
};

/**
 * Returns every slide bound to the given `layout`. Useful for "find
 * all slides using the `Title and Content` layout" workflows; pair
 * with `findSlideLayout(pres, name)` to look up the layout by name.
 */
export const getSlidesByLayout = (
  pres: PresentationData,
  layout: SlideLayoutData,
): ReadonlyArray<SlideData> => {
  const target = layout[LAYOUT_PART_NAME];
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    const found = getSlideLayout(slide);
    if (found !== null && found[LAYOUT_PART_NAME] === target) out.push(slide);
  }
  return out;
};

/**
 * Returns the first slide whose title-placeholder text equals
 * `title` exactly. Returns `null` when no slide matches. Different
 * from `findSlideByText`, which matches any visible text — this one
 * is title-placeholder-scoped. Accepts a literal string (exact
 * equality) or a `RegExp` for pattern matches.
 */
export const findSlideByTitle = (
  pres: PresentationData,
  title: string | RegExp,
): SlideData | null => {
  for (const slide of getSlides(pres)) {
    const t = getSlideTitle(slide);
    if (t === null) continue;
    const hit = typeof title === 'string' ? t === title : title.test(t);
    if (hit) return slide;
  }
  return null;
};

/**
 * Returns the slide whose package part name equals `partName`
 * (typically `/ppt/slides/slideN.xml`), or `null` when no such
 * slide exists. Useful for callers that walk validator output,
 * `listPackageParts`, or any other low-level path-keyed API and
 * need the matching `SlideData` handle.
 */
export const findSlideByPartName = (pres: PresentationData, partName: string): SlideData | null => {
  for (const slide of getSlides(pres)) {
    if (slide[SLIDE_PART_NAME] === partName) return slide;
  }
  return null;
};

/**
 * Returns the package part name (e.g. `/ppt/slides/slide3.xml`) of
 * `slide`. Useful for surfacing slides in validator output, error
 * messages, or any path-keyed sidebar UI.
 */
export const getSlidePartName = (slide: SlideData): string => slide[SLIDE_PART_NAME];

/**
 * Returns the slide's current XML body as a string. Re-serializes
 * from the typed AST, so this reflects any pending edits that
 * haven't been written back to the package yet.
 *
 * Intended for diagnostics: dumping into bug reports, asserting
 * structure in tests, snapshotting before / after a transformation.
 * Do NOT parse this back yourself — round-trip safely through
 * `loadPresentation` / `savePresentation`.
 */
export const getSlideXmlString = (slide: SlideData): string => serializeXml(slide[SLIDE_DOCUMENT]);

/**
 * Returns the first slide whose concatenated visible text contains
 * `needle` (substring; case-sensitive). Pass a `RegExp` to test
 * against the slide's text body instead. Returns `null` when no
 * slide matches.
 */
export const findSlideByText = (
  pres: PresentationData,
  needle: string | RegExp,
): SlideData | null => {
  for (const slide of getSlides(pres)) {
    const text = slideText(slide[SLIDE_PART]);
    if (typeof needle === 'string' ? text.includes(needle) : needle.test(text)) {
      return slide;
    }
  }
  return null;
};

/**
 * Every slide whose concatenated visible text contains `needle`
 * (substring) or matches the given `RegExp`.
 */
export const findSlidesByText = (
  pres: PresentationData,
  needle: string | RegExp,
): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    const text = slideText(slide[SLIDE_PART]);
    if (typeof needle === 'string' ? text.includes(needle) : needle.test(text)) {
      out.push(slide);
    }
  }
  return out;
};

/**
 * Every slide whose speaker notes contain `needle` (substring) or
 * match the given `RegExp`. Sibling of `findSlidesByText` — useful
 * for surfacing reviewer-facing annotations stored in notes (e.g.
 * "every slide whose notes say TODO").
 *
 * Slides without notes are skipped.
 */
export const findSlidesByNotes = (
  pres: PresentationData,
  needle: string | RegExp,
): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    const notes = getSlideNotes(slide);
    if (notes === null || notes.length === 0) continue;
    if (typeof needle === 'string' ? notes.includes(needle) : needle.test(notes)) {
      out.push(slide);
    }
  }
  return out;
};

/**
 * Returns every slide where `needle` appears in **either** the
 * visible text or the speaker notes. Each slide is reported once,
 * in document order. Sibling of `findSlidesByText` /
 * `findSlidesByNotes` — useful for a generic "search the deck"
 * UX where the caller doesn't care which surface matched.
 */
export const searchSlides = (
  pres: PresentationData,
  needle: string | RegExp,
): ReadonlyArray<SlideData> => {
  const matchesString = (haystack: string): boolean =>
    typeof needle === 'string' ? haystack.includes(needle) : needle.test(haystack);
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (matchesString(slideText(slide[SLIDE_PART]))) {
      out.push(slide);
      continue;
    }
    const notes = getSlideNotes(slide);
    if (notes !== null && notes.length > 0 && matchesString(notes)) out.push(slide);
  }
  return out;
};

/**
 * Replaces every occurrence of `from` in every slide's speaker
 * notes with `to`. Returns the number of slides updated. Sibling
 * of `replaceTextInPresentation` for the reviewer-annotation pass.
 *
 * Slides without notes are skipped. For string needles, replacement
 * is global; for RegExp needles, the global flag is forced.
 */
export const replaceTextInNotes = (
  pres: PresentationData,
  from: string | RegExp,
  to: string,
): number => {
  let n = 0;
  for (const slide of getSlides(pres)) {
    if (replaceTextInSlideNotes(slide, from, to)) n++;
  }
  return n;
};

/**
 * Slide-scoped sibling of `replaceTextInNotes`. Replaces every
 * occurrence of `from` in the slide's speaker notes with `to`.
 * Returns `true` if the notes changed, `false` otherwise (no notes
 * present, or no match).
 */
export const replaceTextInSlideNotes = (
  slide: SlideData,
  from: string | RegExp,
  to: string,
): boolean => {
  const notes = getSlideNotes(slide);
  if (notes === null || notes.length === 0) return false;
  let next: string;
  if (typeof from === 'string') {
    if (!notes.includes(from)) return false;
    next = notes.split(from).join(to);
  } else {
    const re = from.global ? from : new RegExp(from.source, `${from.flags}g`);
    if (!re.test(notes)) return false;
    re.lastIndex = 0;
    next = notes.replace(re, to);
  }
  if (next === notes) return false;
  setSlideNotes(slide, next);
  return true;
};

/**
 * Every shape on every slide, paired with its slide and the slide's
 * 0-based index. Useful for cross-deck audits ("find every picture",
 * "list shape names anywhere") without writing the nested-loop
 * boilerplate yourself.
 */
export interface AllShapesEntry {
  readonly slide: SlideData;
  readonly slideIndex: number;
  readonly shape: SlideShapeData;
}

export const getAllShapes = (pres: PresentationData): ReadonlyArray<AllShapesEntry> => {
  const out: AllShapesEntry[] = [];
  const slides = getSlides(pres);
  for (let i = 0; i < slides.length; i++) {
    const slide = slides[i]!;
    for (const shape of getSlideShapes(slide)) {
      out.push({ slide, slideIndex: i, shape });
    }
  }
  return out;
};

// ---------------------------------------------------------------------------
// Slide visibility — `<p:sld show="0">` hides a slide from the slideshow
// without removing it from the deck. `show="1"` (or omission) is visible.

const ATTR_SHOW = qname('', 'show', '');

/**
 * Returns `true` when the slide carries `show="0"` on its root
 * `<p:sld>` element. PowerPoint hides such slides from the slideshow
 * but keeps them in the editing surface.
 */
export const isSlideHidden = (slide: SlideData): boolean => {
  const show = getAttrValue(slide[SLIDE_DOCUMENT].root, ATTR_SHOW);
  return show === '0';
};

/**
 * Returns `true` when the slide carries a `<p:timing>` block — i.e.,
 * has at least one authored animation effect. Per-slide complement to
 * `getPresentationSummary().hasAnimations`, which only reports a
 * deck-wide flag.
 */
export const slideHasAnimations = (slide: SlideData): boolean => {
  return slide[SLIDE_DOCUMENT].root.children.some(
    (c) => c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'timing',
  );
};

/**
 * Toggles the slide's visibility in the slideshow. Hiding adds
 * `show="0"`; showing removes the attribute (PowerPoint treats absence
 * as the default `show="1"`).
 */
export const setSlideHidden = (slide: SlideData, hidden: boolean): void => {
  const root = slide[SLIDE_DOCUMENT].root;
  root.attrs = root.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'show'),
  );
  if (hidden) root.attrs.push(attr(ATTR_SHOW, '0'));
  commitSlideData(slide);
};

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

/**
 * Replaces every occurrence of `from` in every slide's text with `to`.
 * `from` may be a string (treated as a literal) or a `RegExp`. Returns
 * the number of `<a:t>` elements mutated across the whole deck.
 *
 * Use this for the broad "rename product X to Y" pattern; for
 * `{{token}}` style substitutions, prefer
 * `replaceTokensInPresentation` which is more explicit about intent.
 */
export const replaceTextInPresentation = (
  pres: PresentationData,
  from: string | RegExp,
  to: string,
): number => {
  const pkg = pres[INTERNAL_PACKAGE];
  let count = 0;
  for (const part of pkg.parts) {
    if (part.contentType !== SLIDE_CONTENT_TYPE) continue;
    const doc = parseXml(decode(part.data));
    const n = replaceTextInTree(doc.root, from, to);
    if (n > 0) {
      part.data = encode(serializeXml(doc));
      count += n;
    }
  }
  pres._slidesCache = null;
  return count;
};

/**
 * Replaces every occurrence of `from` in the slide's text with `to`.
 * Returns the number of `<a:t>` elements mutated on this slide.
 */
export const replaceTextInSlide = (slide: SlideData, from: string | RegExp, to: string): number => {
  const n = replaceTextInTree(slide[SLIDE_DOCUMENT].root, from, to);
  if (n > 0) {
    commitSlideData(slide);
    refreshSlideData(slide);
  }
  return n;
};
