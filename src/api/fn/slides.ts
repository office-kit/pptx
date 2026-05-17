// Slide-level operations: enumeration, queries, search, replace,
// visibility, and deck manipulation (add/remove/sort/move/duplicate/
// import/merge).

import { replaceTextInTree, replaceTokensInTree } from '../../internal/drawingml/index.ts';
import {
  basename,
  emptyRels,
  nextRelId,
  type PartName,
  partName,
  relsPartNameFor,
  resolveTarget,
} from '../../internal/opc/index.ts';
import type { OpcPackage } from '../../internal/parts/index.ts';
import {
  REL_TYPES,
  buildSlideFromLayout,
  readPresentationPart,
  readSlidePart,
  slideText,
} from '../../internal/presentationml/index.ts';
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
  SLIDE_PART,
  SLIDE_PART_NAME,
  SLIDE_SHAPES,
  type SlideData,
  type SlideLayoutData,
  type SlideShapeData,
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
  SLIDE_LAYOUT_CONTENT_TYPE,
  commitSlideData,
  decode,
  encode,
  refreshSlideData,
  setOpcDefault,
} from './_helpers.ts';
import {
  findSlideLayoutByType,
  getSlideLayoutPlaceholders,
  getSlideLayoutType,
  getSlideLayouts,
} from './package.ts';
import { findSlidePlaceholder, getSlideLayout, getSlideShapes } from './shapes.ts';
import { getSlideNotes, setSlideNotes } from './features.ts';
import { getSlideTitle, setSlideBody, setSlideTitle } from './embedded.ts';

// Build a SlideData handle from a part's bytes. Lives here because
// getSlides is the only call site that loads from cold; other callers
// rebuild the shape view via _helpers.rebuildShapesFromDocument.
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
 * from `findSlideByText`, which matches any visible text via
 * substring / regex — this one is a strict equality match against
 * the title placeholder only.
 */
export const findSlideByTitle = (pres: PresentationData, title: string): SlideData | null => {
  for (const slide of getSlides(pres)) {
    if (getSlideTitle(slide) === title) return slide;
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
  pres._slidesCache = null;
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
 * Imports a slide from another presentation into `targetPres`. The
 * slide's part bytes are copied verbatim; image rels are followed and
 * the linked media is copied into the target package with fresh part
 * names. The new slide is bound to the supplied `targetLayout` so it
 * still renders without the original deck's layouts.
 *
 * Limitations (v1):
 *
 *   - Only `image` rels are copied across. Other rels (charts, embedded
 *     workbooks, oleObjects, comments) are dropped from the imported
 *     slide. A diagnostic message is appended for each dropped rel.
 *   - Hyperlinks (external URLs) are preserved.
 *   - Slide → notesSlide is dropped (notes don't follow imports).
 *
 * Returns the new `SlideData` appended to `targetPres`.
 */
export const importSlide = (
  targetPres: PresentationData,
  sourceSlide: SlideData,
  targetLayout: SlideLayoutData,
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

  // Copy the source slide bytes verbatim.
  targetPkg.addPart(newSlidePartName, sourcePart.contentType, new Uint8Array(sourcePart.data));

  // Build the new slide's rels:
  //   - one slideLayout pointing at the supplied target layout
  //   - one image rel per source image (with media imported)
  //   - external hyperlink rels copied verbatim
  const newRels = emptyRels();
  const layoutPartName = targetLayout[LAYOUT_PART_NAME];
  if (targetPkg.getPart(layoutPartName) === null) {
    throw new Error(`importSlide: layout ${layoutPartName} not in target package`);
  }
  newRels.items.push({
    id: 'rId1',
    type: REL_TYPES.slideLayout,
    target: `../slideLayouts/${basename(layoutPartName)}`,
    targetMode: 'Internal',
  });

  // Map from source rId → new rId so we can rewrite blip references later
  // (skipped in v1; we just preserve original rId values when possible).
  if (sourceRels !== null) {
    for (const rel of sourceRels.items) {
      if (rel.type === REL_TYPES.slideLayout) continue; // handled above
      if (rel.type === REL_TYPES.notesSlide) continue;
      if (rel.type === REL_TYPES.image && rel.targetMode === 'Internal') {
        // Copy the media part across with a fresh name.
        const mediaName = rel.target.startsWith('/')
          ? partName(rel.target)
          : resolveTarget(sourcePartName, rel.target);
        const mediaPart = sourcePkg.getPart(mediaName);
        if (!mediaPart) continue;
        const dotIdx = mediaName.lastIndexOf('.');
        const extension = dotIdx >= 0 ? mediaName.slice(dotIdx + 1) : 'bin';
        let nextN = 1;
        const re = /^\/ppt\/media\/image(\d+)\./;
        for (const p of targetPkg.parts) {
          const m = p.name.match(re);
          if (m?.[1] !== undefined) {
            const n = Number.parseInt(m[1], 10);
            if (Number.isFinite(n) && n >= nextN) nextN = n + 1;
          }
        }
        const newMediaName = partName(`/ppt/media/image${nextN}.${extension}`);
        setOpcDefault(targetPkg, extension.toLowerCase(), mediaPart.contentType);
        targetPkg.addPart(newMediaName, mediaPart.contentType, new Uint8Array(mediaPart.data));
        newRels.items.push({
          id: rel.id,
          type: REL_TYPES.image,
          target: `../media/image${nextN}.${extension}`,
          targetMode: 'Internal',
        });
        continue;
      }
      if (rel.type === REL_TYPES.hyperlink) {
        newRels.items.push({ ...rel });
        continue;
      }
      // Other internal rels (chart/oleObject/etc) are dropped in v1.
    }
  }
  targetPkg.setRels(newSlidePartName, newRels);

  // presentation → slide rel + sldIdLst entry.
  const presRels = targetPkg.getRels(PRES_PART_NAME) ?? emptyRels();
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

  targetPres._slidesCache = null;
  const slides = getSlides(targetPres);
  const last = slides[slides.length - 1];
  if (!last) throw new Error('importSlide: post-condition failed');
  return last;
};

/**
 * Appends every slide from `sourcePres` into `targetPres`, in source
 * order. Built on top of `importSlide`: media is propagated, charts
 * are dropped (not yet supported across decks), and the slide's
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
