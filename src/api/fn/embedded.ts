// Embedded-object operations: comments, hyperlinks, click actions,
// charts, tables, validation, package introspection, image effects,
// animations, and slide-title sugar.

import {
  applyAlignmentToAllParagraphs,
  applyFormatToAllRuns,
  clearFill as clearFillImpl,
  getPictureEmbedRId,
  setSolidFill,
  setTextBody,
  type TextFormat,
  type ParagraphAlignment,
} from '../../internal/drawingml/index.ts';
import type { Emu } from '../units.ts';
import {
  type ImageFormat,
  type PartName,
  basename,
  contentTypeForFormat,
  detectImageFormat,
  emptyRels,
  extensionForFormat,
  nextRelId,
  partName,
  resolveTarget,
} from '../../internal/opc/index.ts';
import type { OpcPackage } from '../../internal/parts/index.ts';
import {
  REL_TYPES,
  type AnimationEffect,
  type AnimationOptions,
  type CommentAuthor,
  type CommentPosition,
  type ShapeKind,
  type SlideComment,
  type SlideLayoutType,
  buildCommentAuthorListDoc,
  buildCommentListDoc,
  buildSingleEffectTiming,
  buildTableCell,
  buildTableRow,
  readCommentAuthorList,
  readCommentList,
} from '../../internal/presentationml/index.ts';
import {
  type IssueSeverity,
  type ValidationIssue,
  validatePresentationPackage,
} from '../../internal/validator/index.ts';
import {
  type ChartKind,
  type ChartSeries,
  type ChartSpec,
  buildChartSpaceDoc,
  buildEmbeddedXlsx,
  readChartSpec,
} from '../../internal/chartml/index.ts';
import {
  NS,
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
} from '../../internal/xml/index.ts';
import {
  CELL_COL,
  CELL_ELEMENT,
  CELL_ROW,
  CELL_TABLE,
  COMMENT_SLIDE,
  COMMENT_SNAPSHOT,
  INTERNAL_PACKAGE,
  type PresentationData,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  SLIDE_DOCUMENT,
  SLIDE_PART_NAME,
  SLIDE_SHAPES,
  type SlideCommentData,
  type SlideData,
  type SlideShapeData,
  type TableCellData,
} from '../_internal-symbols.ts';
import {
  PRES_PART_NAME,
  appendAndReturnNewShape,
  commitAndRefresh,
  commitSlideData,
  decode,
  encode,
  nextShapeId,
  refreshSlideData,
  setOpcDefault,
} from './_helpers.ts';
import {
  getPresentationTheme,
  getSlideLayoutName,
  getSlideLayoutType,
  getSlideLayouts,
  getSlideSections,
} from './package.ts';
import { getSlides, isSlideHidden } from './slides.ts';
import {
  findSlidePlaceholder,
  getShapeBounds,
  getShapeFlip,
  getSlideLayout,
  resolveDrawingColor,
  setShapeText,
  shapesOverlap,
} from './shapes.ts';
import { getSlideSize } from './features.ts';

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

const writeCommentsForSlide = (slide: SlideData, comments: ReadonlyArray<SlideComment>): void => {
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

const asCommentData = (
  slide: SlideData,
  snap: SlideComment,
  author: CommentAuthor,
): SlideCommentData => ({
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
 * Returns every slide that has at least one comment by the given
 * author name. Sibling of `findCommentsByAuthor` (which returns
 * the comments themselves). Case-sensitive equality.
 */
export const findSlidesWithCommentsByAuthor = (
  pres: PresentationData,
  authorName: string,
): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    const hit = getSlideComments(slide).some((c) => c.author.name === authorName);
    if (hit) out.push(slide);
  }
  return out;
};

/**
 * Returns every distinct author who has at least one comment
 * anywhere in the deck. Deduplicates by author id; preserves
 * first-seen order. Differs from `getCommentAuthors(pres)`, which
 * surfaces every author registered in `commentAuthors.xml` even
 * when no comments reference them.
 */
export const getPresentationCommenters = (pres: PresentationData): ReadonlyArray<CommentAuthor> => {
  const seen = new Set<number>();
  const out: CommentAuthor[] = [];
  for (const slide of getSlides(pres)) {
    for (const c of getSlideComments(slide)) {
      if (seen.has(c.author.id)) continue;
      seen.add(c.author.id);
      out.push(c.author);
    }
  }
  return out;
};

/**
 * Looks up a `CommentAuthor` from `commentAuthors.xml` by display
 * name (case-sensitive equality). Returns `null` when no author has
 * that name. Sibling of `findCommentsByAuthor` — the latter returns
 * the matching comments; this returns the author handle for
 * downstream metadata reads (id, initials, color).
 */
export const findCommentAuthorByName = (
  pres: PresentationData,
  authorName: string,
): CommentAuthor | null => {
  for (const a of getCommentAuthors(pres)) {
    if (a.name === authorName) return a;
  }
  return null;
};

/**
 * Returns every comment whose author name matches `authorName`
 * exactly, across every slide in the deck. Useful for reviewer-
 * specific filters ("show me all of Alice's notes").
 *
 * Case-sensitive equality. Use `findShapesByText` (with a slide
 * predicate of your choosing) for fuzzier text matching against
 * comment bodies.
 */
export const findCommentsByAuthor = (
  pres: PresentationData,
  authorName: string,
): ReadonlyArray<SlideCommentData> => {
  const out: SlideCommentData[] = [];
  for (const slide of getSlides(pres)) {
    for (const c of getSlideComments(slide)) {
      if (c.author.name === authorName) out.push(c);
    }
  }
  return out;
};

/**
 * Returns every comment whose text matches `needle` (substring or
 * `RegExp`) across the whole deck. Sibling of `findCommentsByAuthor`
 * — useful for "find every comment that mentions X" reviewer flows.
 */
export const findCommentsByText = (
  pres: PresentationData,
  needle: string | RegExp,
): ReadonlyArray<SlideCommentData> => {
  const out: SlideCommentData[] = [];
  for (const slide of getSlides(pres)) {
    for (const c of getSlideComments(slide)) {
      const text = c[COMMENT_SNAPSHOT].text;
      const hit = typeof needle === 'string' ? text.includes(needle) : needle.test(text);
      if (hit) out.push(c);
    }
  }
  return out;
};

/**
 * Returns every comment whose `@dt` timestamp is **strictly after**
 * `since` (an ISO-8601 string or `Date`). Comments missing a `dt`
 * are skipped. Sibling of `findCommentsByText` / `findCommentsByAuthor`
 * — useful for "what's new since my last review?" surfaces.
 */
export const findCommentsAfter = (
  pres: PresentationData,
  since: string | Date,
): ReadonlyArray<SlideCommentData> => {
  const threshold = typeof since === 'string' ? Date.parse(since) : since.getTime();
  if (Number.isNaN(threshold)) return [];
  const out: SlideCommentData[] = [];
  for (const slide of getSlides(pres)) {
    for (const c of getSlideComments(slide)) {
      const dt = c[COMMENT_SNAPSHOT].dt;
      if (dt === null) continue;
      const t = Date.parse(dt);
      if (Number.isNaN(t)) continue;
      if (t > threshold) out.push(c);
    }
  }
  return out;
};

/**
 * Returns every comment whose `@dt` timestamp is **strictly before**
 * `until` (an ISO-8601 string or `Date`). Comments missing a `dt`
 * are skipped. Sibling of `findCommentsAfter`.
 */
export const findCommentsBefore = (
  pres: PresentationData,
  until: string | Date,
): ReadonlyArray<SlideCommentData> => {
  const threshold = typeof until === 'string' ? Date.parse(until) : until.getTime();
  if (Number.isNaN(threshold)) return [];
  const out: SlideCommentData[] = [];
  for (const slide of getSlides(pres)) {
    for (const c of getSlideComments(slide)) {
      const dt = c[COMMENT_SNAPSHOT].dt;
      if (dt === null) continue;
      const t = Date.parse(dt);
      if (Number.isNaN(t)) continue;
      if (t < threshold) out.push(c);
    }
  }
  return out;
};

const datedComments = (
  pres: PresentationData,
): ReadonlyArray<{ comment: SlideCommentData; t: number }> => {
  const out: { comment: SlideCommentData; t: number }[] = [];
  for (const slide of getSlides(pres)) {
    for (const c of getSlideComments(slide)) {
      const dt = c[COMMENT_SNAPSHOT].dt;
      if (dt === null) continue;
      const t = Date.parse(dt);
      if (!Number.isNaN(t)) out.push({ comment: c, t });
    }
  }
  return out;
};

/**
 * Returns every comment carrying a parseable `@dt`, sorted oldest
 * to newest. Comments without a date are omitted. Use `.at(0)` for
 * the oldest and `.at(-1)` for the newest.
 */
export const getCommentsSortedByDate = (
  pres: PresentationData,
): ReadonlyArray<SlideCommentData> => {
  const dated = [...datedComments(pres)];
  dated.sort((a, b) => a.t - b.t);
  return dated.map((d) => d.comment);
};

/**
 * Returns the distinct authors who commented on this slide, in
 * first-seen order. Dedupes by author id. Sibling of
 * `getPresentationCommenters` for a slide-scoped reviewer roster.
 */
export const getSlideCommentAuthors = (slide: SlideData): ReadonlyArray<CommentAuthor> => {
  const seen = new Set<number>();
  const out: CommentAuthor[] = [];
  for (const c of getSlideComments(slide)) {
    if (seen.has(c.author.id)) continue;
    seen.add(c.author.id);
    out.push(c.author);
  }
  return out;
};

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

/**
 * Strips every comment from every slide in the deck. Returns the
 * number of comments removed. Built on `writeCommentsForSlide`,
 * so each slide's modern comment part is rewritten with an empty
 * list. The `commentAuthors.xml` registry is left intact for any
 * caller that still needs author identity.
 *
 * Useful as a sanitizer before sharing a draft externally — pairs
 * with `clearAllSlideNotes` for a "remove reviewer chatter" pass.
 */
export const clearAllSlideComments = (pres: PresentationData): number => {
  let n = 0;
  for (const slide of getSlides(pres)) n += clearSlideComments(slide);
  return n;
};

/**
 * Slide-scoped sibling of `clearAllSlideComments`. Removes every
 * comment on the given slide and returns the number removed.
 */
export const clearSlideComments = (slide: SlideData): number => {
  const comments = loadCommentsForSlide(slide);
  if (comments.length === 0) return 0;
  writeCommentsForSlide(slide, []);
  return comments.length;
};

// Accessors over CommentAuthor / SlideCommentData for tree-shake convenience.

export const getCommentAuthor = (comment: SlideCommentData): CommentAuthor => comment.author;
export const getCommentText = (comment: SlideCommentData): string => comment[COMMENT_SNAPSHOT].text;
export const getCommentDate = (comment: SlideCommentData): string | null =>
  comment[COMMENT_SNAPSHOT].dt;
export const getCommentPosition = (comment: SlideCommentData): CommentPosition | null =>
  comment[COMMENT_SNAPSHOT].position;

/**
 * Returns the slide that owns this comment. Counterpart to
 * `getShapeSlide(shape)` — handy when filtering across the deck
 * with `findCommentsByAuthor` and the caller needs to know which
 * slide each hit came from.
 */
export const getCommentSlide = (comment: SlideCommentData): SlideData => comment[COMMENT_SLIDE];

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
    throw new Error(
      `setShapeImage only works on picture shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const format = options.format ?? detectImageFormat(bytes);
  if (format === null) {
    throw new Error(
      'setShapeImage: could not detect image format. Pass options.format explicitly.',
    );
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

export const NAME_HLINK_CLICK_FN = qname('a', 'hlinkClick', NS.dml);

// cNvPr lives at different paths depending on shape kind. Returns null
// for kinds we don't know how to navigate yet (groups, etc.).
export const findCNvPr = (shape: SlideShapeData): XmlElement | null => {
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
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        c.name.localName === 'hlinkClick'
      ),
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
 * Reads the click action attached to the shape's cNvPr, or `null` if
 * none. Mirrors `setShapeClickAction`:
 *
 *   - `{ kind: 'url', url }`     — `hyperlink` rel + targetMode=External
 *   - `{ kind: 'slide', slide }` — `slide` rel + `ppaction://hlinksldjump`
 *   - `{ kind: 'nextSlide' | 'prevSlide' | 'firstSlide' | 'lastSlide' }`
 *     — preset show-navigation `ppaction`.
 *
 * For `kind: 'slide'`, the matching slide is resolved by part name.
 * Returns `null` for unknown `ppaction` strings.
 */
export const getShapeClickAction = (shape: SlideShapeData): ShapeClickAction | null => {
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) return null;
  const hlink = firstChildElement(cNvPr, NAME_HLINK_CLICK_FN);
  if (!hlink) return null;
  const action = getAttrValue(hlink, qname('', 'action', ''));
  const rId = getAttrValue(hlink, qname('r', 'id', NS.officeDocRels));

  if (action === 'ppaction://hlinkshowjump?jump=nextslide') return { kind: 'nextSlide' };
  if (action === 'ppaction://hlinkshowjump?jump=previousslide') return { kind: 'prevSlide' };
  if (action === 'ppaction://hlinkshowjump?jump=firstslide') return { kind: 'firstSlide' };
  if (action === 'ppaction://hlinkshowjump?jump=lastslide') return { kind: 'lastSlide' };

  if (rId !== null && rId !== '') {
    const slide = shape[SHAPE_SLIDE];
    const pkg = slide[INTERNAL_PACKAGE];
    const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
    if (!rels) return null;
    const rel = rels.items.find((r) => r.id === rId);
    if (!rel) return null;

    if (action === 'ppaction://hlinksldjump' && rel.type === REL_TYPES.slide) {
      // Resolve to the SlideData of the target slide.
      const targetPartName = rel.target.startsWith('/')
        ? partName(rel.target)
        : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
      const pres: PresentationData = { [INTERNAL_PACKAGE]: pkg, _slidesCache: null };
      for (const candidate of getSlides(pres)) {
        if (candidate[SLIDE_PART_NAME] === targetPartName) {
          return { kind: 'slide', slide: candidate };
        }
      }
      return null;
    }
    if (rel.type === REL_TYPES.hyperlink && rel.targetMode === 'External') {
      return { kind: 'url', url: rel.target };
    }
  }
  return null;
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

const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
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

// ---------------------------------------------------------------------------
// Table cell access.
//
// `addSlideTable` builds the cell tree under `<a:graphic>/<a:graphicData>/
// <a:tbl>`. These helpers let callers reach individual `<a:tc>` cells to
// re-fill text, paint backgrounds, or align contents without rebuilding
// the table.

export type { TableCellData };

const NAME_A_GRAPHIC_TBL = qname('a', 'graphic', NS.dml);
const NAME_A_GRAPHIC_DATA_TBL = qname('a', 'graphicData', NS.dml);
const NAME_A_TBL = qname('a', 'tbl', NS.dml);
const NAME_A_TC_PR = qname('a', 'tcPr', NS.dml);
const NAME_A_TX_BODY_TBL = qname('a', 'txBody', NS.dml);

const findTblElement = (shape: SlideShapeData): XmlElement | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'graphicFrame') return null;
  const graphic = firstChildElement(shape[SHAPE_ELEMENT], NAME_A_GRAPHIC_TBL);
  if (!graphic) return null;
  const graphicData = firstChildElement(graphic, NAME_A_GRAPHIC_DATA_TBL);
  if (!graphicData) return null;
  return firstChildElement(graphicData, NAME_A_TBL);
};

/**
 * `true` when the shape is a `<p:graphicFrame>` wrapping `<a:tbl>` —
 * i.e. a table. Sharper than `getShapeKind(shape) === 'graphicFrame'`,
 * which also matches charts and SmartArt frames.
 */
export const isTableShape = (shape: SlideShapeData): boolean => findTblElement(shape) !== null;

/**
 * `true` when the shape is a `<p:graphicFrame>` wrapping a chart
 * reference (`<c:chart>`). Charts, tables, and SmartArt all share
 * the graphic-frame kind; this predicate filters down to charts only.
 */
export const isChartShape = (shape: SlideShapeData): boolean =>
  resolveChartPartName(shape[SHAPE_SLIDE], shape) !== null;

const NAME_A_GRID_COL = qname('a', 'gridCol', NS.dml);
const ATTR_W_TBL = qname('', 'w', '');
const ATTR_H_TBL = qname('', 'h', '');

const tableRows = (tbl: XmlElement): XmlElement[] =>
  tbl.children.filter(
    (c): c is XmlElement =>
      c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'tr',
  );

const rowCells = (tr: XmlElement): XmlElement[] =>
  tr.children.filter(
    (c): c is XmlElement =>
      c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'tc',
  );

const buildCellHandle = (
  table: SlideShapeData,
  tc: XmlElement,
  row: number,
  col: number,
): TableCellData => ({
  [CELL_TABLE]: table,
  [CELL_ELEMENT]: tc,
  [CELL_ROW]: row,
  [CELL_COL]: col,
});

/**
 * Returns a 2D array of cell handles for the given table shape, in
 * row-major order. Throws if the shape isn't a table graphic frame.
 */
export const getTableCells = (
  table: SlideShapeData,
): ReadonlyArray<ReadonlyArray<TableCellData>> => {
  const tbl = findTblElement(table);
  if (!tbl) throw new Error('getTableCells: shape is not a table graphic frame');
  const rows = tableRows(tbl);
  return rows.map((tr, rowIdx) =>
    rowCells(tr).map((tc, colIdx) => buildCellHandle(table, tc, rowIdx, colIdx)),
  );
};

/**
 * Reads the table-style GUID from `<a:tbl><a:tblPr><a:tableStyleId>`.
 * PowerPoint uses GUIDs to reference built-in table styles
 * (`{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}` = "Medium Style 2 -
 * Accent 1", etc.) and theme-local styles. Returns `null` when the
 * table doesn't reference one (uses the slide's default style).
 */
export const getTableStyleId = (table: SlideShapeData): string | null => {
  const tbl = findTblElement(table);
  if (!tbl) return null;
  const tblPr = firstChildElement(tbl, qname('a', 'tblPr', NS.dml));
  if (!tblPr) return null;
  const idEl = firstChildElement(tblPr, qname('a', 'tableStyleId', NS.dml));
  if (!idEl) return null;
  let acc = '';
  for (const c of idEl.children) {
    if (c.kind === 'text' || c.kind === 'cdata') acc += c.data;
  }
  return acc.trim() || null;
};

/**
 * Reads the boolean style flags off `<a:tblPr>` — which header /
 * footer rows + columns are banded or emphasised. Mirrors the
 * `firstRow` / `bandRow` flags exposed by `addSlideTable`.
 *
 * Per ECMA-376 §17.18.95 / §21.1.3.15:
 *
 *   - `firstRow` — header row is styled differently.
 *   - `lastRow` — total row is styled differently.
 *   - `firstCol` — first column is styled differently.
 *   - `lastCol` — last column is styled differently.
 *   - `bandRow` — alternating row shading.
 *   - `bandCol` — alternating column shading.
 *
 * Renderers use these to switch on the corresponding table style
 * variant. Returns all-`false` for tables that don't author `<a:tblPr>`.
 */
export const getTableStyleFlags = (
  table: SlideShapeData,
): {
  firstRow: boolean;
  lastRow: boolean;
  firstCol: boolean;
  lastCol: boolean;
  bandRow: boolean;
  bandCol: boolean;
} => {
  const empty = {
    firstRow: false,
    lastRow: false,
    firstCol: false,
    lastCol: false,
    bandRow: false,
    bandCol: false,
  };
  const tbl = findTblElement(table);
  if (!tbl) return empty;
  const tblPr = firstChildElement(tbl, qname('a', 'tblPr', NS.dml));
  if (!tblPr) return empty;
  const readBool = (attr: string): boolean => {
    const v = getAttrValue(tblPr, qname('', attr, ''));
    return v === '1' || v === 'true';
  };
  return {
    firstRow: readBool('firstRow'),
    lastRow: readBool('lastRow'),
    firstCol: readBool('firstCol'),
    lastCol: readBool('lastCol'),
    bandRow: readBool('bandRow'),
    bandCol: readBool('bandCol'),
  };
};

/**
 * Returns the table's row + column counts. Throws when the shape
 * isn't a table graphic frame.
 */
export const getTableDimensions = (
  table: SlideShapeData,
): { readonly rows: number; readonly cols: number } => {
  const tbl = findTblElement(table);
  if (!tbl) throw new Error('getTableDimensions: shape is not a table graphic frame');
  const rows = tableRows(tbl);
  const cols = rows[0] !== undefined ? rowCells(rows[0]).length : 0;
  return { rows: rows.length, cols };
};

/**
 * Returns each column's width in EMU, in left-to-right order, as
 * declared on `<a:tblGrid>/<a:gridCol w="...">`. Missing or
 * unparseable widths default to 0. Throws when the shape isn't a
 * table graphic frame.
 */
export const getTableColumnWidths = (table: SlideShapeData): ReadonlyArray<Emu> => {
  const tbl = findTblElement(table);
  if (!tbl) throw new Error('getTableColumnWidths: shape is not a table graphic frame');
  const grid = firstChildElement(tbl, qname('a', 'tblGrid', NS.dml));
  if (!grid) return [];
  const out: Emu[] = [];
  for (const col of allChildElements(grid, NAME_A_GRID_COL)) {
    const v = getAttrValue(col, ATTR_W_TBL);
    const n = v !== null ? Number.parseInt(v, 10) : 0;
    out.push((Number.isFinite(n) ? n : 0) as Emu);
  }
  return out;
};

/**
 * Returns the table's nominal `(width, height)` derived from
 * summing the `<a:gridCol w>` and `<a:tr h>` attributes (both in
 * EMU). Useful for layout pipelines that want to know how big a
 * table really is without dereferencing the shape's `<a:xfrm>`.
 *
 * Throws when the shape isn't a table graphic frame.
 */
export const getTableSize = (
  table: SlideShapeData,
): { readonly width: Emu; readonly height: Emu } => {
  const widths = getTableColumnWidths(table);
  const heights = getTableRowHeights(table);
  const width = widths.reduce((sum, w) => sum + w, 0) as Emu;
  const height = heights.reduce((sum, h) => sum + h, 0) as Emu;
  return { width, height };
};

/**
 * Returns each row's height in EMU, in top-to-bottom order, from
 * `<a:tr h="...">`. Missing or unparseable heights default to 0.
 * Throws when the shape isn't a table graphic frame.
 */
export const getTableRowHeights = (table: SlideShapeData): ReadonlyArray<Emu> => {
  const tbl = findTblElement(table);
  if (!tbl) throw new Error('getTableRowHeights: shape is not a table graphic frame');
  const out: Emu[] = [];
  for (const tr of tableRows(tbl)) {
    const v = getAttrValue(tr, ATTR_H_TBL);
    const n = v !== null ? Number.parseInt(v, 10) : 0;
    out.push((Number.isFinite(n) ? n : 0) as Emu);
  }
  return out;
};

/**
 * Sets a single column's width on the table grid. Throws on
 * out-of-range column indices or non-table shapes. The total table
 * width is not auto-adjusted — callers are responsible for keeping
 * the sum consistent with the table's `<a:xfrm>` extent if PowerPoint
 * is to render the table without clipping.
 */
export const setTableColumnWidth = (table: SlideShapeData, col: number, width: Emu): void => {
  const tbl = requireTbl(table);
  const grid = firstChildElement(tbl, qname('a', 'tblGrid', NS.dml));
  if (!grid) throw new Error('table has no <a:tblGrid>');
  const cols = allChildElements(grid, NAME_A_GRID_COL);
  const target = cols[col];
  if (!target) throw new RangeError(`table column ${col} out of range (have ${cols.length})`);
  target.attrs = target.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'w'),
  );
  target.attrs.push(attr(ATTR_W_TBL, String(Math.round(width))));
  commitSlideData(table[SHAPE_SLIDE]);
  refreshSlideData(table[SHAPE_SLIDE]);
};

/**
 * Sets a single row's height. Throws on out-of-range row indices or
 * non-table shapes. As with `setTableColumnWidth`, the table's
 * `<a:xfrm>` extent is left to the caller.
 */
export const setTableRowHeight = (table: SlideShapeData, row: number, height: Emu): void => {
  const tbl = requireTbl(table);
  const rows = tableRows(tbl);
  const target = rows[row];
  if (!target) throw new RangeError(`table row ${row} out of range (have ${rows.length})`);
  target.attrs = target.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'h'),
  );
  target.attrs.push(attr(ATTR_H_TBL, String(Math.round(height))));
  commitSlideData(table[SHAPE_SLIDE]);
  refreshSlideData(table[SHAPE_SLIDE]);
};

/**
 * Returns the cell at `(row, col)`. Throws on out-of-range coordinates
 * or non-table shapes.
 */
export const getTableCell = (table: SlideShapeData, row: number, col: number): TableCellData => {
  const cells = getTableCells(table);
  const r = cells[row];
  if (!r) throw new RangeError(`table row ${row} out of range (have ${cells.length})`);
  const c = r[col];
  if (!c) throw new RangeError(`table column ${col} out of range in row ${row} (have ${r.length})`);
  return c;
};

const commitTableCell = (cell: TableCellData): void => {
  const shape = cell[CELL_TABLE];
  commitSlideData(shape[SHAPE_SLIDE]);
  refreshSlideData(shape[SHAPE_SLIDE]);
};

const ensureCellTxBody = (cell: TableCellData): XmlElement => {
  const tc = cell[CELL_ELEMENT];
  let txBody = firstChildElement(tc, NAME_A_TX_BODY_TBL);
  if (txBody === null) {
    txBody = elem(NAME_A_TX_BODY_TBL);
    // bodyPr lstStyle a:p — keep the canonical ordering.
    txBody.children.push(elem(qname('a', 'bodyPr', NS.dml)));
    txBody.children.push(elem(qname('a', 'lstStyle', NS.dml)));
    // Insert before <a:tcPr> per the schema.
    const tcPrIdx = tc.children.findIndex(
      (c) => c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'tcPr',
    );
    if (tcPrIdx >= 0) tc.children.splice(tcPrIdx, 0, txBody);
    else tc.children.push(txBody);
  }
  return txBody;
};

const ensureCellTcPr = (cell: TableCellData): XmlElement => {
  const tc = cell[CELL_ELEMENT];
  let tcPr = firstChildElement(tc, NAME_A_TC_PR);
  if (tcPr === null) {
    tcPr = elem(NAME_A_TC_PR);
    // <a:tcPr> is the LAST child of <a:tc>.
    tc.children.push(tcPr);
  }
  return tcPr;
};

/** Replaces a cell's text. `\n` starts a new paragraph. */
export const setTableCellText = (cell: TableCellData, text: string): void => {
  const txBody = ensureCellTxBody(cell);
  setTextBody(txBody, text);
  commitTableCell(cell);
};

/**
 * Reads the cell's merge / span attributes:
 *
 *   - `gridSpan` — number of columns this cell spans (≥1, default 1).
 *   - `rowSpan` — number of rows this cell spans (≥1, default 1).
 *   - `hMerge` — `true` when this cell is the right half of a horizontal
 *     span (it's painted by an earlier cell with `gridSpan > 1`).
 *   - `vMerge` — `true` when this cell is the bottom half of a vertical
 *     span (it's painted by an earlier cell with `rowSpan > 1`).
 *
 * Renderers should skip painting cells where `hMerge` or `vMerge` is
 * true; those cells exist only so the row/column grid stays consistent.
 */
export const getTableCellSpan = (
  cell: TableCellData,
): { gridSpan: number; rowSpan: number; hMerge: boolean; vMerge: boolean } => {
  const el = cell[CELL_ELEMENT];
  const gs = getAttrValue(el, qname('', 'gridSpan', ''));
  const rs = getAttrValue(el, qname('', 'rowSpan', ''));
  const hm = getAttrValue(el, qname('', 'hMerge', ''));
  const vm = getAttrValue(el, qname('', 'vMerge', ''));
  const parseSpan = (v: string | null): number => {
    if (v === null) return 1;
    const n = Number.parseInt(v, 10);
    return Number.isFinite(n) && n > 0 ? n : 1;
  };
  return {
    gridSpan: parseSpan(gs),
    rowSpan: parseSpan(rs),
    hMerge: hm === '1' || hm === 'true',
    vMerge: vm === '1' || vm === 'true',
  };
};

/**
 * One side of a cell's border, as read from `<a:tcPr><a:ln{L|R|T|B}>`.
 * `widthEmu` is the line width in EMU; `color` is `#RRGGBB` or `null`
 * when the border is inherited / un-styled.
 */
export interface TableCellBorder {
  readonly color: string | null;
  readonly widthEmu: number | null;
  readonly dash: string | null;
}

export interface TableCellBorders {
  readonly left: TableCellBorder | null;
  readonly right: TableCellBorder | null;
  readonly top: TableCellBorder | null;
  readonly bottom: TableCellBorder | null;
  readonly tlToBr: TableCellBorder | null;
  readonly blToTr: TableCellBorder | null;
}

/**
 * Reads the per-side borders on a cell. Returns `null` for sides with
 * no explicit `<a:ln{Side}>` element (those inherit from the table
 * style / theme). All four cardinal sides plus the two diagonals
 * (TL→BR, BL→TR) are surfaced because real templates use them all.
 */
export const getTableCellBorders = (
  pres: PresentationData,
  cell: TableCellData,
): TableCellBorders => {
  const tcPr = firstChildElement(cell[CELL_ELEMENT], NAME_A_TC_PR);
  const theme = getPresentationTheme(pres);
  const empty: TableCellBorders = {
    left: null,
    right: null,
    top: null,
    bottom: null,
    tlToBr: null,
    blToTr: null,
  };
  if (!tcPr) return empty;
  const readLn = (local: string): TableCellBorder | null => {
    const ln = firstChildElement(tcPr, qname('a', local, NS.dml));
    if (!ln) return null;
    const w = getAttrValue(ln, qname('', 'w', ''));
    const widthEmu = w !== null ? Number.parseInt(w, 10) : null;
    let color: string | null = null;
    const solid = firstChildElement(ln, qname('a', 'solidFill', NS.dml));
    if (solid) {
      for (const c of solid.children) {
        if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
        color = resolveDrawingColor(c, theme);
        break;
      }
    }
    const prstDash = firstChildElement(ln, qname('a', 'prstDash', NS.dml));
    const dash = prstDash ? getAttrValue(prstDash, qname('', 'val', '')) : null;
    return { color, widthEmu, dash };
  };
  return {
    left: readLn('lnL'),
    right: readLn('lnR'),
    top: readLn('lnT'),
    bottom: readLn('lnB'),
    tlToBr: readLn('lnTlToBr'),
    blToTr: readLn('lnBlToTr'),
  };
};

/**
 * Reads the cell's text direction (`<a:tcPr vert="…"/>`) — same tokens
 * as `getShapeTextDirection`. Returns `null` for the default
 * horizontal direction.
 *
 * Vertical column headers in tables almost always emit `<a:tcPr
 * vert="vert270"/>` or `"eaVert"` so the header label reads bottom-to-
 * top alongside its column.
 */
export const getTableCellTextDirection = (
  cell: TableCellData,
): 'vert' | 'vert270' | 'wordArtVert' | 'eaVert' | 'mongolianVert' | 'wordArtVertRtl' | null => {
  const tcPr = firstChildElement(cell[CELL_ELEMENT], NAME_A_TC_PR);
  if (!tcPr) return null;
  const v = getAttrValue(tcPr, qname('', 'vert', ''));
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

/**
 * Reads the cell's vertical text anchor (`<a:tcPr anchor="t|ctr|b"/>`)
 * — `'top'`, `'center'`, `'bottom'`, or `null` for the default
 * (`ctr` / center per the schema).
 */
export const getTableCellAnchor = (cell: TableCellData): 'top' | 'center' | 'bottom' | null => {
  const tcPr = firstChildElement(cell[CELL_ELEMENT], NAME_A_TC_PR);
  if (!tcPr) return null;
  const v = getAttrValue(tcPr, qname('', 'anchor', ''));
  if (v === 't') return 'top';
  if (v === 'ctr') return 'center';
  if (v === 'b') return 'bottom';
  return null;
};

/**
 * Reads the cell's inset margins (`<a:tcPr marL marR marT marB>`) in
 * EMU. Each side is `null` when the cell doesn't author it (renderers
 * should fall back to PowerPoint's defaults — 91440 EMU / 0.1 inch
 * for the horizontal margins, 45720 EMU for the vertical).
 */
export const getTableCellMargins = (
  cell: TableCellData,
): { left: number | null; right: number | null; top: number | null; bottom: number | null } => {
  const tcPr = firstChildElement(cell[CELL_ELEMENT], NAME_A_TC_PR);
  const empty = { left: null, right: null, top: null, bottom: null };
  if (!tcPr) return empty;
  const read = (name: string): number | null => {
    const v = getAttrValue(tcPr, qname('', name, ''));
    if (v === null) return null;
    const n = Number.parseInt(v, 10);
    return Number.isFinite(n) ? n : null;
  };
  return {
    left: read('marL'),
    right: read('marR'),
    top: read('marT'),
    bottom: read('marB'),
  };
};

/** Reads the cell's plain text (paragraphs joined with `\n`). */
export const getTableCellText = (cell: TableCellData): string => {
  const txBody = firstChildElement(cell[CELL_ELEMENT], NAME_A_TX_BODY_TBL);
  if (!txBody) return '';
  const lines: string[] = [];
  for (const p of txBody.children) {
    if (p.kind !== 'element' || p.name.namespaceURI !== NS.dml || p.name.localName !== 'p')
      continue;
    let line = '';
    for (const r of p.children) {
      if (r.kind !== 'element' || r.name.namespaceURI !== NS.dml || r.name.localName !== 'r')
        continue;
      const tEl = firstChildElement(r, qname('a', 't', NS.dml));
      if (!tEl) continue;
      for (const child of tEl.children) {
        if (child.kind === 'text' || child.kind === 'cdata') line += child.data;
      }
    }
    lines.push(line);
  }
  return lines.join('\n');
};

/** Sets a solid background color on a cell (`<a:tcPr><a:solidFill>`). */
export const setTableCellFill = (cell: TableCellData, color: string): void => {
  const tcPr = ensureCellTcPr(cell);
  setSolidFill(tcPr, color);
  commitTableCell(cell);
};

/** Removes any background fill from a cell. */
export const clearTableCellFill = (cell: TableCellData): void => {
  const tcPr = ensureCellTcPr(cell);
  clearFillImpl(tcPr);
  commitTableCell(cell);
};

/**
 * Reads the cell's solid background color. Returns `#RRGGBB` for
 * `<a:srgbClr>`, `scheme:<token>` for `<a:schemeClr>`, or `null` when
 * the cell has no fill, no `<a:tcPr>`, or the fill is non-solid
 * (gradient / pattern / image).
 */
export const getTableCellFill = (cell: TableCellData): string | null => {
  const tcPr = firstChildElement(cell[CELL_ELEMENT], NAME_A_TC_PR);
  if (!tcPr) return null;
  const solid = firstChildElement(tcPr, qname('a', 'solidFill', NS.dml));
  if (!solid) return null;
  const srgb = firstChildElement(solid, qname('a', 'srgbClr', NS.dml));
  if (srgb) {
    const v = getAttrValue(srgb, qname('', 'val', ''));
    if (v) return `#${v.toUpperCase()}`;
  }
  const scheme = firstChildElement(solid, qname('a', 'schemeClr', NS.dml));
  if (scheme) {
    const v = getAttrValue(scheme, qname('', 'val', ''));
    if (v) return `scheme:${v}`;
  }
  return null;
};

/** Applies a TextFormat to every run in the cell's text. */
export const setTableCellTextFormat = (cell: TableCellData, format: TextFormat): void => {
  const txBody = ensureCellTxBody(cell);
  applyFormatToAllRuns(txBody, format);
  commitTableCell(cell);
};

/** Sets horizontal alignment on every paragraph in the cell. */
export const setTableCellAlignment = (cell: TableCellData, align: ParagraphAlignment): void => {
  const txBody = ensureCellTxBody(cell);
  applyAlignmentToAllParagraphs(txBody, align);
  commitTableCell(cell);
};

/** Zero-based (row, col) of the cell. */
export const getTableCellPosition = (cell: TableCellData): { row: number; col: number } => ({
  row: cell[CELL_ROW],
  col: cell[CELL_COL],
});

/**
 * Reads the horizontal alignment from the cell's first paragraph
 * (`l`, `ctr`, `r`, `just`, `dist`, `justLow`, `thaiDist`). Returns
 * `null` when the cell has no `<a:txBody>` or its first paragraph
 * has no explicit `algn` attribute (PowerPoint then defaults to `l`).
 */
export const getTableCellAlignment = (cell: TableCellData): ParagraphAlignment | null => {
  const txBody = firstChildElement(cell[CELL_ELEMENT], NAME_A_TX_BODY_TBL);
  if (!txBody) return null;
  for (const p of txBody.children) {
    if (p.kind !== 'element' || p.name.namespaceURI !== NS.dml || p.name.localName !== 'p')
      continue;
    const pPr = firstChildElement(p, qname('a', 'pPr', NS.dml));
    if (!pPr) return null;
    const v = getAttrValue(pPr, qname('', 'algn', ''));
    return (v as ParagraphAlignment | null) ?? null;
  }
  return null;
};

const requireTbl = (table: SlideShapeData): XmlElement => {
  const tbl = findTblElement(table);
  if (!tbl) throw new Error('table shape is not a table graphic frame');
  return tbl;
};

const tableColumnCount = (tbl: XmlElement): number => {
  const grid = firstChildElement(tbl, qname('a', 'tblGrid', NS.dml));
  if (!grid) return 0;
  return allChildElements(grid, NAME_A_GRID_COL).length;
};

const rowDefaultHeight = (tbl: XmlElement): number => {
  // Use the average height of existing rows, or 370000 (≈ 0.4in) as a
  // sane default when the table has no rows yet.
  const rows = tableRows(tbl);
  if (rows.length === 0) return 370000;
  let sum = 0;
  let count = 0;
  for (const r of rows) {
    const h = getAttrValue(r, ATTR_H_TBL);
    if (h !== null) {
      const n = Number.parseInt(h, 10);
      if (Number.isFinite(n)) {
        sum += n;
        count++;
      }
    }
  }
  return count > 0 ? Math.round(sum / count) : 370000;
};

/**
 * Inserts a row into the table. `atIndex` is 0-based; `undefined`
 * appends at the end. `cells` supplies cell values; missing cells
 * become blank, extras are dropped. The row's height matches the
 * average of existing rows (or a 0.4in default for empty tables).
 */
export const insertTableRow = (
  table: SlideShapeData,
  atIndex?: number,
  cells: ReadonlyArray<string> = [],
): void => {
  const tbl = requireTbl(table);
  const colCount = tableColumnCount(tbl);
  const padded: string[] = [];
  for (let i = 0; i < colCount; i++) padded.push(cells[i] ?? '');
  const row = buildTableRow(padded, rowDefaultHeight(tbl));

  const rows = tableRows(tbl);
  const insertAt =
    atIndex !== undefined ? Math.max(0, Math.min(atIndex, rows.length)) : rows.length;
  if (insertAt === rows.length) {
    tbl.children.push(row);
  } else {
    const target = rows[insertAt]!;
    const idx = tbl.children.indexOf(target);
    tbl.children.splice(idx, 0, row);
  }
  commitSlideData(table[SHAPE_SLIDE]);
  refreshSlideData(table[SHAPE_SLIDE]);
};

/** Removes the row at `atIndex` from the table. Throws on out-of-range. */
export const removeTableRow = (table: SlideShapeData, atIndex: number): void => {
  const tbl = requireTbl(table);
  const rows = tableRows(tbl);
  if (atIndex < 0 || atIndex >= rows.length) {
    throw new RangeError(`removeTableRow: index ${atIndex} out of range (have ${rows.length})`);
  }
  const target = rows[atIndex]!;
  tbl.children = tbl.children.filter((c) => c !== target);
  commitSlideData(table[SHAPE_SLIDE]);
  refreshSlideData(table[SHAPE_SLIDE]);
};

/**
 * Inserts a column into the table. `atIndex` defaults to the end.
 * `widthEmu` defaults to the average of existing column widths (or
 * 914400 = 1in if the table has no columns). Existing rows get a new
 * blank cell at `atIndex`.
 */
export const insertTableColumn = (
  table: SlideShapeData,
  atIndex?: number,
  widthEmu?: number,
): void => {
  const tbl = requireTbl(table);
  const grid = firstChildElement(tbl, qname('a', 'tblGrid', NS.dml));
  if (!grid) throw new Error('table is missing <a:tblGrid>');
  const cols = allChildElements(grid, NAME_A_GRID_COL);
  const insertAt =
    atIndex !== undefined ? Math.max(0, Math.min(atIndex, cols.length)) : cols.length;

  // Default width: average of existing widths.
  let defaultWidth = widthEmu;
  if (defaultWidth === undefined) {
    let sum = 0;
    let count = 0;
    for (const col of cols) {
      const w = getAttrValue(col, ATTR_W_TBL);
      if (w !== null) {
        const n = Number.parseInt(w, 10);
        if (Number.isFinite(n)) {
          sum += n;
          count++;
        }
      }
    }
    defaultWidth = count > 0 ? Math.round(sum / count) : 914400;
  }
  const newCol = elem(NAME_A_GRID_COL, { attrs: [attr(ATTR_W_TBL, String(defaultWidth))] });
  if (insertAt === cols.length) {
    grid.children.push(newCol);
  } else {
    const target = cols[insertAt]!;
    const idx = grid.children.indexOf(target);
    grid.children.splice(idx, 0, newCol);
  }

  // Add a blank <a:tc> at the same column index in every row.
  for (const tr of tableRows(tbl)) {
    const tcs = rowCells(tr);
    const newCell = buildTableCell('');
    if (insertAt >= tcs.length) {
      tr.children.push(newCell);
    } else {
      const target = tcs[insertAt]!;
      const idx = tr.children.indexOf(target);
      tr.children.splice(idx, 0, newCell);
    }
  }

  commitSlideData(table[SHAPE_SLIDE]);
  refreshSlideData(table[SHAPE_SLIDE]);
};

/** Removes the column at `atIndex` (and the corresponding cell in every row). */
export const removeTableColumn = (table: SlideShapeData, atIndex: number): void => {
  const tbl = requireTbl(table);
  const grid = firstChildElement(tbl, qname('a', 'tblGrid', NS.dml));
  if (!grid) throw new Error('table is missing <a:tblGrid>');
  const cols = allChildElements(grid, NAME_A_GRID_COL);
  if (atIndex < 0 || atIndex >= cols.length) {
    throw new RangeError(`removeTableColumn: index ${atIndex} out of range (have ${cols.length})`);
  }
  const targetCol = cols[atIndex]!;
  grid.children = grid.children.filter((c) => c !== targetCol);
  for (const tr of tableRows(tbl)) {
    const tcs = rowCells(tr);
    if (atIndex < tcs.length) {
      tr.children = tr.children.filter((c) => c !== tcs[atIndex]);
    }
  }
  commitSlideData(table[SHAPE_SLIDE]);
  refreshSlideData(table[SHAPE_SLIDE]);
};

/**
 * A chart sitting on a slide. `shape` is the `<p:graphicFrame>`
 * wrapper; `spec` is the chart definition parsed from the linked
 * `/ppt/charts/chart{N}.xml` part. `null` `spec` means the chart uses
 * a kind we don't model (callers can fall through to pass-through).
 */
export interface SlideChartData {
  readonly shape: SlideShapeData;
  readonly spec: ChartSpec | null;
}

const NAME_A_GRAPHIC_FN = qname('a', 'graphic', NS.dml);
const NAME_A_GRAPHIC_DATA_FN = qname('a', 'graphicData', NS.dml);
const NAME_C_CHART_FN = qname('c', 'chart', NS.chart);

/**
 * Resolves the chart part backing a graphic-frame shape, or `null` if
 * the shape isn't a chart wrapper.
 */
const resolveChartPartName = (
  slide: SlideData,
  shape: SlideShapeData,
): { partName: PartName; rId: string } | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'graphicFrame') return null;
  const graphic = firstChildElement(shape[SHAPE_ELEMENT], NAME_A_GRAPHIC_FN);
  if (!graphic) return null;
  const graphicData = firstChildElement(graphic, NAME_A_GRAPHIC_DATA_FN);
  if (!graphicData) return null;
  const chartRef = firstChildElement(graphicData, NAME_C_CHART_FN);
  if (!chartRef) return null;
  const rId = getAttrValue(chartRef, qname('r', 'id', NS.officeDocRels));
  if (rId === null) return null;
  const slideRels = slide[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]);
  if (!slideRels) return null;
  const rel = slideRels.items.find((r) => r.id === rId);
  if (!rel) return null;
  const partNameValue = rel.target.startsWith('/')
    ? partName(rel.target)
    : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
  return { partName: partNameValue, rId };
};

/**
 * Replaces the chart definition on an existing graphic-frame chart
 * shape. Updates the inline `<c:strCache>` / `<c:numCache>` blocks so
 * PowerPoint renders the new data without opening the embedded
 * workbook. The shape's geometry (position / size / rotation) is
 * preserved verbatim.
 *
 * The embedded xlsx is re-written too — it's what the "Edit data"
 * affordance opens. The previous workbook is replaced wholesale (no
 * attempt to preserve styles a user added through Excel).
 *
 * Pass any `ChartSpec`, including a different `kind` from the
 * original; this acts as "change my column chart to a line chart with
 * fresh data."
 */
export const setChartSpec = (chart: SlideChartData, spec: ChartSpec): void => {
  const slide = chart.shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  const resolved = resolveChartPartName(slide, chart.shape);
  if (!resolved) {
    throw new Error('setChartSpec: shape is not a chart graphic frame');
  }

  // Rewrite the chart XML.
  const doc = buildChartSpaceDoc(spec);
  const chartBytes = encode(serializeXml(doc));
  const chartPart = pkg.getPart(resolved.partName);
  if (!chartPart) {
    throw new Error(`setChartSpec: chart part ${resolved.partName} not found`);
  }
  chartPart.data = chartBytes;

  // Rewrite the embedded xlsx (the part chart→package rel points at).
  // Reuse the existing rel; fall back to creating a fresh xlsx part if
  // the chart had no package rel (unusual for charts we authored).
  const chartRels = pkg.getRels(resolved.partName);
  if (chartRels) {
    const xlsxRel = chartRels.items.find((r) => r.type === REL_TYPES.package);
    if (xlsxRel) {
      const xlsxName = xlsxRel.target.startsWith('/')
        ? partName(xlsxRel.target)
        : resolveTarget(resolved.partName, xlsxRel.target);
      const xlsxPart = pkg.getPart(xlsxName);
      const rows = spec.categories.map((label, i) => ({
        label,
        values: spec.series.map((s) => s.values[i] ?? null),
      }));
      const xlsxBytes = buildEmbeddedXlsx(
        spec.series.map((s) => s.name),
        rows,
      );
      if (xlsxPart) {
        xlsxPart.data = xlsxBytes;
      }
    }
  }
};

/**
 * Returns every chart on the slide, with its `ChartSpec` parsed from
 * the linked chart part. Skips graphic frames that don't carry a
 * `<c:chart>` reference (e.g. tables, diagrams).
 */
/**
 * For a graphic-frame shape that wraps a chart, returns the parsed
 * `ChartSpec`. Returns `null` when the shape isn't a chart wrapper
 * or the chart uses a kind we don't model (e.g. surface, radar).
 *
 * Convenience over `getSlideCharts(...).find((c) => c.shape === shape)`
 * when the caller already has the shape in hand (e.g. iterating
 * `getSlideShapes`).
 */
/**
 * Convenience over `getShapeChartSpec(shape)?.kind ?? null`. Returns
 * the chart's `ChartKind` ('bar', 'line', 'pie', …) when the shape
 * is a chart wrapper, or `null` for non-charts and charts whose
 * kind isn't modeled yet.
 */
export const getShapeChartKind = (shape: SlideShapeData): ChartKind | null => {
  const spec = getShapeChartSpec(shape);
  return spec === null ? null : spec.kind;
};

/**
 * Returns the categories axis labels of a chart shape, or `null`
 * if the shape isn't a chart wrapper or its kind isn't modeled.
 * Convenience over `getShapeChartSpec(shape)?.categories ?? null`.
 */
export const getShapeChartCategories = (shape: SlideShapeData): ReadonlyArray<string> | null => {
  const spec = getShapeChartSpec(shape);
  return spec === null ? null : spec.categories;
};

/**
 * Returns the chart's series-name list (in spec order). `null`
 * when the shape isn't a chart wrapper or the kind isn't modeled.
 */
export const getShapeChartSeriesNames = (shape: SlideShapeData): ReadonlyArray<string> | null => {
  const spec = getShapeChartSpec(shape);
  return spec === null ? null : spec.series.map((s) => s.name);
};

/**
 * Returns the values for the named series on a chart shape, or
 * `null` when the shape isn't a chart, the kind isn't modeled, or
 * no series matches `seriesName`.
 */
export const getShapeChartSeriesValues = (
  shape: SlideShapeData,
  seriesName: string,
): ReadonlyArray<number | null> | null => {
  const spec = getShapeChartSpec(shape);
  if (spec === null) return null;
  const series = spec.series.find((s) => s.name === seriesName);
  return series ? series.values : null;
};

export const getShapeChartSpec = (shape: SlideShapeData): ChartSpec | null => {
  const slide = shape[SHAPE_SLIDE];
  const resolved = resolveChartPartName(slide, shape);
  if (!resolved) return null;
  const part = slide[INTERNAL_PACKAGE].getPart(resolved.partName);
  if (!part) return null;
  try {
    const root = parseXml(decode(part.data)).root;
    return readChartSpec(root);
  } catch {
    return null;
  }
};

/**
 * Returns every chart on the slide that carries a series whose name
 * equals `seriesName` exactly. Useful for "find the revenue chart"
 * patterns where chart kind alone isn't unique. Skips charts whose
 * kind isn't modeled.
 */
export const findChartsBySeriesName = (
  slide: SlideData,
  seriesName: string,
): ReadonlyArray<SlideChartData> => {
  const out: SlideChartData[] = [];
  for (const chart of getSlideCharts(slide)) {
    if (chart.spec === null) continue;
    if (chart.spec.series.some((s) => s.name === seriesName)) out.push(chart);
  }
  return out;
};

/**
 * Returns the first chart on the slide whose parsed `kind` matches
 * `kind` (e.g. `'bar'`, `'line'`, `'pie'`). Returns `null` when no
 * chart on the slide has that kind, or when every chart on the slide
 * uses a kind this version doesn't yet model.
 */
export const findChartByKind = (slide: SlideData, kind: ChartKind): SlideChartData | null => {
  for (const chart of getSlideCharts(slide)) {
    if (chart.spec !== null && chart.spec.kind === kind) return chart;
  }
  return null;
};

export const getSlideCharts = (slide: SlideData): ReadonlyArray<SlideChartData> => {
  const pkg = slide[INTERNAL_PACKAGE];
  const out: SlideChartData[] = [];
  for (const shape of slide[SLIDE_SHAPES]) {
    const resolved = resolveChartPartName(slide, shape);
    if (!resolved) continue;
    const chartPart = pkg.getPart(resolved.partName);
    if (!chartPart) continue;
    let spec: ChartSpec | null;
    try {
      const root = parseXml(decode(chartPart.data)).root;
      spec = readChartSpec(root);
    } catch {
      spec = null;
    }
    out.push({ shape, spec });
  }
  return out;
};

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
// Package introspection escape hatches.
//
// `_internalPackageOf` is the heavy escape hatch for hot-path power
// users; these two helpers cover the 80% case (just enumerate parts
// or read a single part's bytes) without exposing the OpcPackage
// class.

/**
 * Power-user escape hatch. Returns the underlying `OpcPackage`
 * backing `pres`. Use this when you need to manipulate parts /
 * rels directly. Most callers should use the typed helpers above
 * (`listPackageParts`, `readPackagePart`, `getMediaParts`, etc.).
 *
 * @internal — used by `pptx-kit/node` to mount fs-backed helpers.
 */
export const _internalPackageOf = (pres: PresentationData): OpcPackage => pres[INTERNAL_PACKAGE];

/** One entry in the package's parts list. */
export interface PackagePartInfo {
  readonly name: string;
  readonly contentType: string;
  readonly byteLength: number;
}

/**
 * Enumerates every OPC part in the package. Useful for advanced
 * inspection (e.g. "what parts does this template carry?") without
 * dropping to `_internalPackageOf`.
 */
export const listPackageParts = (pres: PresentationData): ReadonlyArray<PackagePartInfo> =>
  pres[INTERNAL_PACKAGE].parts.map((p) => ({
    name: p.name,
    contentType: p.contentType,
    byteLength: p.data.byteLength,
  }));

/**
 * Reads a single OPC part's bytes by part name (e.g.
 * `'/ppt/slides/slide1.xml'`). Returns `null` when no such part
 * exists. The returned `Uint8Array` is a live view into the
 * package — DO NOT mutate it. Use this for read-only inspection
 * (e.g. parsing custom extension parts).
 */
export const readPackagePart = (pres: PresentationData, name: string): Uint8Array | null => {
  const part = pres[INTERNAL_PACKAGE].parts.find((p) => p.name === name);
  return part?.data ?? null;
};

/** A media (image / video / audio) part embedded in the package. */
export interface MediaPart {
  readonly name: string;
  readonly contentType: string;
  readonly data: Uint8Array;
}

/**
 * Returns the total size of the package's parts in bytes
 * (uncompressed). Useful for storage estimation, quota checks,
 * and "how big is this deck before save?" diagnostics. The
 * actual `savePresentation` output is typically smaller after
 * DEFLATE; this is an upper bound.
 */
export const getPackageSize = (pres: PresentationData): number => {
  let total = 0;
  for (const part of pres[INTERNAL_PACKAGE].parts) total += part.data.byteLength;
  return total;
};

/**
 * Returns every `/ppt/media/...` part in the package. Useful for
 * audit / export workflows — e.g. "extract every embedded image."
 */
export const getMediaParts = (pres: PresentationData): ReadonlyArray<MediaPart> => {
  const out: MediaPart[] = [];
  for (const p of pres[INTERNAL_PACKAGE].parts) {
    if (p.name.startsWith('/ppt/media/')) {
      out.push({ name: p.name, contentType: p.contentType, data: p.data });
    }
  }
  return out;
};

/**
 * Returns every media part NOT referenced by any rels in the
 * package — the set `compactPackage` would remove. Non-destructive;
 * the caller decides whether to delete.
 *
 * Useful for audit UIs that want to surface bloat before cleaning,
 * and for "is this asset still used?" checks.
 */
export const getOrphanMediaPartNames = (pres: PresentationData): ReadonlyArray<string> => {
  const pkg = pres[INTERNAL_PACKAGE];
  const referenced = new Set<string>();
  const resolve = (sourcePart: string, target: string): string => {
    if (target.startsWith('/')) return target;
    const dir = sourcePart.split('/').slice(0, -1);
    const segments: string[] = [];
    for (const seg of [...dir, ...target.split('/')]) {
      if (seg === '..') segments.pop();
      else if (seg !== '.' && seg.length > 0) segments.push(seg);
    }
    return `/${segments.join('/')}`;
  };
  for (const part of pkg.parts) {
    if (!part.name.endsWith('.rels')) continue;
    // /ppt/slides/_rels/slide1.xml.rels → /ppt/slides/slide1.xml
    const m = part.name.match(/^(.*)\/_rels\/(.+)\.rels$/);
    let sourceName: string;
    if (part.name === '/_rels/.rels') {
      sourceName = '/';
    } else if (m) {
      sourceName = `${m[1]}/${m[2]}`;
    } else {
      continue;
    }
    const sourceRels = sourceName === '/' ? pkg.rootRels() : pkg.getRels(sourceName as never);
    if (!sourceRels) continue;
    for (const rel of sourceRels.items) {
      if (rel.targetMode === 'External') continue;
      referenced.add(resolve(sourceName, rel.target));
    }
  }
  const out: string[] = [];
  for (const part of pkg.parts) {
    if (!part.name.startsWith('/ppt/media/')) continue;
    if (!referenced.has(part.name)) out.push(part.name);
  }
  return out;
};

/**
 * Returns every media part name the slide's rels reference
 * (typically `/ppt/media/imageN.ext`). Walks the slide's rels
 * graph and resolves each internal target. Useful for "which
 * media files does this slide depend on?" audits.
 */
export const getSlideMediaPartNames = (slide: SlideData): ReadonlyArray<string> => {
  const pkg = slide[INTERNAL_PACKAGE];
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return [];
  const resolve = (sourcePart: string, target: string): string => {
    if (target.startsWith('/')) return target;
    const dir = sourcePart.split('/').slice(0, -1);
    const segments: string[] = [];
    for (const seg of [...dir, ...target.split('/')]) {
      if (seg === '..') segments.pop();
      else if (seg !== '.' && seg.length > 0) segments.push(seg);
    }
    return `/${segments.join('/')}`;
  };
  const out: string[] = [];
  const seen = new Set<string>();
  for (const rel of rels.items) {
    if (rel.targetMode === 'External') continue;
    const resolved = resolve(slide[SLIDE_PART_NAME], rel.target);
    if (!resolved.startsWith('/ppt/media/')) continue;
    if (seen.has(resolved)) continue;
    seen.add(resolved);
    out.push(resolved);
  }
  return out;
};

/**
 * Returns every slide that references the given media part name
 * (typically `/ppt/media/imageN.ext`). Walks each slide's rels and
 * checks whether any internal rel resolves to `mediaPartName`.
 *
 * Useful for image-audit workflows: "before I replace this image,
 * which slides will the change affect?"
 */
export const slidesUsingMediaPart = (
  pres: PresentationData,
  mediaPartName: string,
): ReadonlyArray<SlideData> => {
  const pkg = pres[INTERNAL_PACKAGE];
  const resolve = (sourcePart: string, target: string): string => {
    if (target.startsWith('/')) return target;
    const dir = sourcePart.split('/').slice(0, -1);
    const segments: string[] = [];
    for (const seg of [...dir, ...target.split('/')]) {
      if (seg === '..') segments.pop();
      else if (seg !== '.' && seg.length > 0) segments.push(seg);
    }
    return `/${segments.join('/')}`;
  };
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
    if (!rels) continue;
    const hit = rels.items.some(
      (r) =>
        r.targetMode !== 'External' && resolve(slide[SLIDE_PART_NAME], r.target) === mediaPartName,
    );
    if (hit) out.push(slide);
  }
  return out;
};

/**
 * Removes media parts that no rels graph references. Returns the
 * list of removed part names. Useful after a sequence of slide
 * removals leaves orphan images behind.
 *
 * Only `/ppt/media/...` parts are considered. The check walks every
 * `.rels` part in the package and resolves each internal rel target
 * against its source part name to build the live media set.
 */
export const compactPackage = (
  pres: PresentationData,
): { readonly removed: ReadonlyArray<string> } => {
  const pkg = pres[INTERNAL_PACKAGE];
  const referenced = new Set<string>();

  const resolve = (sourcePart: string, target: string): string => {
    if (target.startsWith('/')) return target;
    const dir = sourcePart.split('/').slice(0, -1);
    const segments: string[] = [];
    for (const seg of [...dir, ...target.split('/')]) {
      if (seg === '..') segments.pop();
      else if (seg !== '.' && seg.length > 0) segments.push(seg);
    }
    return `/${segments.join('/')}`;
  };

  for (const part of pkg.parts) {
    if (!part.name.endsWith('.rels')) continue;
    // /ppt/slides/_rels/slide1.xml.rels → /ppt/slides/slide1.xml
    // /_rels/.rels                       → / (package root)
    let sourceName = part.name.replace('/_rels/', '/').replace(/\.rels$/, '');
    if (sourceName === '/' || sourceName === '') {
      // Root rels — `rel.target` is relative to the package root.
      // We don't need to consult pkg.getRels for it (the only thing it
      // points at that we care about is the presentation.xml, which
      // has its own rels we'll walk). Just parse the part data
      // directly for completeness.
      sourceName = '/';
    }
    const rels = sourceName === '/' ? null : pkg.getRels(partName(sourceName));
    if (!rels) continue;
    for (const rel of rels.items) {
      if (rel.targetMode === 'External') continue;
      referenced.add(resolve(sourceName, rel.target));
    }
  }

  const removed: string[] = [];
  const orphans: string[] = [];
  for (const part of pkg.parts) {
    if (!part.name.startsWith('/ppt/media/')) continue;
    if (!referenced.has(part.name)) orphans.push(part.name);
  }
  for (const name of orphans) {
    pkg.removePart(partName(name));
    removed.push(name);
  }
  return { removed };
};

/**
 * Replaces the bytes of a media part in place. Returns `true` if the
 * part was found and updated, `false` otherwise. The content type is
 * preserved.
 *
 * Useful for the "swap every instance of this logo" workflow — pick
 * the right `partName` via `getMediaParts` and call this once. Every
 * `<a:blip r:embed="…"/>` reference is unaffected because the rels
 * already point at this part name.
 */
export const setMediaPartBytes = (
  pres: PresentationData,
  partName: string,
  bytes: Uint8Array,
): boolean => {
  const part = pres[INTERNAL_PACKAGE].parts.find((p) => p.name === partName);
  if (!part) return false;
  part.data = bytes;
  return true;
};

/**
 * High-level snapshot of the presentation's structure. Useful as a
 * diagnostic checklist when debugging a template or generating audit
 * reports. The numbers reflect what's reachable through the typed
 * API on the current package state.
 */
export interface PresentationSummary {
  readonly slideCount: number;
  readonly hiddenSlideCount: number;
  readonly totalShapes: number;
  readonly shapesByKind: Readonly<Record<ShapeKind, number>>;
  readonly layoutCount: number;
  readonly sectionCount: number;
  readonly partCount: number;
  readonly hasCharts: boolean;
  readonly hasComments: boolean;
  readonly hasAnimations: boolean;
  readonly themeName: string | null;
}

const CHART_CONTENT_TYPE_FN = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const COMMENTS_CONTENT_TYPE_FN =
  'application/vnd.openxmlformats-officedocument.presentationml.comments+xml';

export const getPresentationSummary = (pres: PresentationData): PresentationSummary => {
  const pkg = pres[INTERNAL_PACKAGE];
  const slides = getSlides(pres);
  let hiddenSlideCount = 0;
  let totalShapes = 0;
  const shapesByKind: Record<ShapeKind, number> = {
    shape: 0,
    picture: 0,
    group: 0,
    graphicFrame: 0,
    connector: 0,
  };
  let hasAnimations = false;
  for (const slide of slides) {
    if (isSlideHidden(slide)) hiddenSlideCount++;
    for (const s of slide[SLIDE_SHAPES]) {
      totalShapes++;
      shapesByKind[s[SHAPE_SNAPSHOT].kind]++;
    }
    // <p:timing> presence = at least one animation.
    if (
      !hasAnimations &&
      slide[SLIDE_DOCUMENT].root.children.some(
        (c) =>
          c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'timing',
      )
    ) {
      hasAnimations = true;
    }
  }

  const hasCharts = pkg.parts.some((p) => p.contentType === CHART_CONTENT_TYPE_FN);
  const hasComments = pkg.parts.some((p) => p.contentType === COMMENTS_CONTENT_TYPE_FN);
  const theme = getPresentationTheme(pres);

  return {
    slideCount: slides.length,
    hiddenSlideCount,
    totalShapes,
    shapesByKind,
    layoutCount: getSlideLayouts(pres).length,
    sectionCount: getSlideSections(pres).length,
    partCount: pkg.parts.length,
    hasCharts,
    hasComments,
    hasAnimations,
    themeName: theme?.name ?? null,
  };
};

// ---------------------------------------------------------------------------
// Picture opacity — `<a:alphaModFix>` inside the picture's `<a:blip>`.
//
// `amt` is ECMA-376's ST_PositiveFixedPercentage (0–100000, scale 1/1000
// of a percent). PowerPoint defaults to fully opaque when the element
// is absent. Pass `null` to remove a prior `<a:alphaModFix>`.

const NAME_ALPHA_MOD_FIX_FN = qname('a', 'alphaModFix', NS.dml);
const ATTR_AMT_FN = qname('', 'amt', '');

/**
 * Sets the picture's opacity (0–1 fraction; `1` is fully opaque, `0`
 * fully transparent). Pass `null` to remove an existing opacity
 * override and restore PowerPoint's default behavior.
 *
 * Throws for non-picture shapes and on opacities outside `[0, 1]`.
 */
/**
 * Returns the embedded image bytes for a picture shape, or `null`
 * when the shape isn't a picture or has no `r:embed` reference
 * (external images aren't followed).
 *
 * The returned `Uint8Array` is a live view into the package media
 * part — treat it as read-only; copy if you need an independent
 * buffer.
 */
export const getShapeImageBytes = (shape: SlideShapeData): Uint8Array | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') return null;
  const rEmbed = getPictureEmbedRId(shape[SHAPE_ELEMENT]);
  if (rEmbed === null) return null;
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;
  const rel = rels.items.find((r) => r.id === rEmbed);
  if (!rel || rel.targetMode === 'External') return null;
  const mediaName = rel.target.startsWith('/')
    ? partName(rel.target)
    : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
  const part = pkg.getPart(mediaName);
  return part?.data ?? null;
};

/**
 * `true` when the shape's text body carries any visible characters.
 * Tighter than checking `getShapeText(shape) !== ''` because it
 * doesn't allocate the concatenated string.
 */
export const hasShapeText = (shape: SlideShapeData): boolean => {
  const text = shape[SHAPE_SNAPSHOT].text;
  return typeof text === 'string' && text.length > 0;
};

/**
 * `true` when the shape carries an embedded image — either a
 * `<p:pic>` picture or a `<p:spPr>/<a:blipFill>` image fill on a
 * regular shape. External `r:link` references count too.
 */
export const hasShapeImage = (shape: SlideShapeData): boolean => {
  if (shape[SHAPE_SNAPSHOT].kind === 'picture') {
    return getPictureEmbedRId(shape[SHAPE_ELEMENT]) !== null;
  }
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return false;
  return firstChildElement(spPr, qname('a', 'blipFill', NS.dml)) !== null;
};

/**
 * Returns every shape on the slide that is mirrored — horizontally
 * (`flipH`), vertically (`flipV`), or both.
 */
export const findFlippedShapes = (slide: SlideData): ReadonlyArray<SlideShapeData> =>
  slide[SLIDE_SHAPES].filter((s) => {
    const flip = getShapeFlip(s);
    return flip !== null && (flip.horizontal || flip.vertical);
  });

/**
 * Returns every unordered pair of shapes on the slide whose
 * bounding boxes overlap. Built on `shapesOverlap`. Pairs are
 * returned with `a` strictly preceding `b` in document order, and
 * each pair appears at most once.
 *
 * Useful for layout audits — "do any boxes collide on this slide?"
 * Shapes without `<a:xfrm>` bounds never overlap anything.
 */
export const findOverlappingShapePairs = (
  slide: SlideData,
): ReadonlyArray<readonly [SlideShapeData, SlideShapeData]> => {
  const shapes = slide[SLIDE_SHAPES];
  const out: (readonly [SlideShapeData, SlideShapeData])[] = [];
  for (let i = 0; i < shapes.length; i++) {
    for (let j = i + 1; j < shapes.length; j++) {
      if (shapesOverlap(shapes[i]!, shapes[j]!)) {
        out.push([shapes[i]!, shapes[j]!] as const);
      }
    }
  }
  return out;
};

/**
 * Returns every shape on the slide whose bounding box extends past
 * the slide canvas (`getSlideSize(pres)`). Useful audit helper for
 * catching shapes that PowerPoint will silently render off-screen
 * or clip on export. Shapes without `<a:xfrm>` bounds are skipped.
 *
 * If the presentation has no slide-size declared, every positioned
 * shape is returned (caller can't audit against an absent canvas).
 */
export const findShapesOutsideCanvas = (
  slide: SlideData,
  pres: PresentationData,
): ReadonlyArray<SlideShapeData> => {
  const size = getSlideSize(pres);
  const out: SlideShapeData[] = [];
  for (const shape of slide[SLIDE_SHAPES]) {
    const b = getShapeBounds(shape);
    if (b === null) continue;
    if (size === null) {
      out.push(shape);
      continue;
    }
    if (b.x < 0 || b.y < 0 || b.x + b.w > size.width || b.y + b.h > size.height) {
      out.push(shape);
    }
  }
  return out;
};

/**
 * Every slide whose layout's `cSld@name` matches the given string.
 * Useful for batch operations on slides sharing a layout — for
 * example, restyling every "Title and Content" slide in a deck.
 *
 * Matching is exact (case-sensitive). Slides without a resolved
 * layout are skipped.
 */
export const findSlidesByLayoutName = (
  pres: PresentationData,
  layoutName: string,
): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    const layout = getSlideLayout(slide);
    if (layout !== null && getSlideLayoutName(layout) === layoutName) out.push(slide);
  }
  return out;
};

/**
 * Every slide whose layout `@type` (e.g. `'title'`, `'blank'`,
 * `'obj'`) matches. Sibling of `findSlidesByLayoutName`, but keyed
 * on the OOXML layout-type enum rather than the human-facing name —
 * stable across locales and template providers.
 */
export const findSlidesByLayoutType = (
  pres: PresentationData,
  layoutType: SlideLayoutType | string,
): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    const layout = getSlideLayout(slide);
    if (layout !== null && getSlideLayoutType(layout) === layoutType) out.push(slide);
  }
  return out;
};

/**
 * Returns the package part name (`/ppt/media/imageN.ext`) of
 * whichever image the shape carries — picture (`<p:pic>`) or
 * image-fill (`<a:blipFill>` nested under `<p:spPr>`). Returns
 * `null` when the shape has no embedded image, or the rel points
 * at an external `r:link` rather than an internal target.
 *
 * Useful for addressing the media part directly with
 * `setMediaPartBytes` or `readPackagePart`.
 */
/**
 * Returns the external URL of the picture when its `<a:blip>` carries an
 * `r:link` (external) relationship rather than an `r:embed`. Returns
 * `null` for embedded pictures, non-picture shapes, or when the
 * relationship doesn't resolve.
 *
 * PowerPoint emits `r:link` when the user inserts via "Link to file"
 * instead of "Insert Picture". The bytes live outside the package, so
 * `getShapeImageBytes` can't render them — readers / preview tools
 * should fall back to this URL or a placeholder.
 */
/**
 * Returns `true` when the picture's `<a:blip>` carries `<a:grayscl/>`
 * — PowerPoint's "Color > Grayscale" recolor preset. Renderers can
 * project this onto a CSS `filter: grayscale(100%)` or an SVG
 * `<feColorMatrix>` desaturation.
 */
export const isShapeImageGrayscale = (shape: SlideShapeData): boolean => {
  let blip: XmlElement | null = null;
  if (shape[SHAPE_SNAPSHOT].kind === 'picture') {
    const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
    if (blipFill) blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  } else {
    const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
    if (spPr) {
      const bf = firstChildElement(spPr, qname('a', 'blipFill', NS.dml));
      if (bf) blip = firstChildElement(bf, qname('a', 'blip', NS.dml));
    }
  }
  return blip !== null && firstChildElement(blip, qname('a', 'grayscl', NS.dml)) !== null;
};

/**
 * Returns the threshold of the picture's `<a:blip><a:biLevel thresh="…"/>`
 * effect — PowerPoint's "Color > Black and White" preset. Threshold is
 * a percent (0..100); pixels brighter become white, darker become black.
 * Returns `null` when no biLevel transform is set.
 */
export const getShapeImageBiLevelThreshold = (shape: SlideShapeData): number | null => {
  let blip: XmlElement | null = null;
  if (shape[SHAPE_SNAPSHOT].kind === 'picture') {
    const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
    if (blipFill) blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  } else {
    const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
    if (spPr) {
      const bf = firstChildElement(spPr, qname('a', 'blipFill', NS.dml));
      if (bf) blip = firstChildElement(bf, qname('a', 'blip', NS.dml));
    }
  }
  if (!blip) return null;
  const biLevel = firstChildElement(blip, qname('a', 'biLevel', NS.dml));
  if (!biLevel) return null;
  const t = getAttrValue(biLevel, qname('', 'thresh', ''));
  if (t === null) return null;
  let n = Number.parseFloat(t);
  if (!Number.isFinite(n)) return null;
  if (Math.abs(n) > 1) n = n / 100000;
  return n * 100;
};

/**
 * Reads the picture's duotone color transform from `<a:blip><a:duotone>`.
 * PowerPoint emits two `<a:srgbClr>` (or scheme color) children for a
 * two-color duotone effect — typical "Picture Tools › Recolor".
 * Returns `null` when no duotone is set.
 */
export const getShapeImageDuotone = (
  pres: PresentationData,
  shape: SlideShapeData,
): { firstColor: string | null; secondColor: string | null } | null => {
  let blip: XmlElement | null = null;
  if (shape[SHAPE_SNAPSHOT].kind === 'picture') {
    const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
    if (blipFill) blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  } else {
    const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
    if (spPr) {
      const bf = firstChildElement(spPr, qname('a', 'blipFill', NS.dml));
      if (bf) blip = firstChildElement(bf, qname('a', 'blip', NS.dml));
    }
  }
  if (!blip) return null;
  const duotone = firstChildElement(blip, qname('a', 'duotone', NS.dml));
  if (!duotone) return null;
  const theme = getPresentationTheme(pres);
  const colors: Array<string | null> = [];
  for (const c of duotone.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (
      c.name.localName === 'srgbClr' ||
      c.name.localName === 'schemeClr' ||
      c.name.localName === 'sysClr' ||
      c.name.localName === 'prstClr'
    ) {
      colors.push(resolveDrawingColor(c, theme));
      if (colors.length === 2) break;
    }
  }
  return {
    firstColor: colors[0] ?? null,
    secondColor: colors[1] ?? null,
  };
};

export const getShapeImageLinkUrl = (shape: SlideShapeData): string | null => {
  const slide = shape[SHAPE_SLIDE];
  const rels = slide[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;
  // Find the blip element on either picture or shape-with-image-fill.
  let blip: XmlElement | null = null;
  if (shape[SHAPE_SNAPSHOT].kind === 'picture') {
    const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
    if (blipFill) blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  } else {
    const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
    if (spPr) {
      const bf = firstChildElement(spPr, qname('a', 'blipFill', NS.dml));
      if (bf) blip = firstChildElement(bf, qname('a', 'blip', NS.dml));
    }
  }
  if (!blip) return null;
  const rLink = getAttrValue(blip, qname('r', 'link', NS.officeDocRels));
  if (!rLink) return null;
  const rel = rels.items.find((r) => r.id === rLink);
  if (!rel || rel.targetMode !== 'External') return null;
  return rel.target;
};

export const getShapeImagePartName = (shape: SlideShapeData): string | null => {
  const slide = shape[SHAPE_SLIDE];
  const rels = slide[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;

  const resolve = (rEmbed: string | null): string | null => {
    if (rEmbed === null) return null;
    const rel = rels.items.find((r) => r.id === rEmbed);
    if (!rel || rel.targetMode === 'External') return null;
    const name = rel.target.startsWith('/')
      ? partName(rel.target)
      : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
    return name;
  };

  // Picture shape: <p:pic><p:blipFill><a:blip r:embed="..."/>.
  if (shape[SHAPE_SNAPSHOT].kind === 'picture') {
    const rEmbed = getPictureEmbedRId(shape[SHAPE_ELEMENT]);
    return resolve(rEmbed);
  }

  // Other shapes with image fill: <p:spPr><a:blipFill><a:blip r:embed="..."/>.
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const blipFill = firstChildElement(spPr, qname('a', 'blipFill', NS.dml));
  if (!blipFill) return null;
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) return null;
  return resolve(getAttrValue(blip, qname('r', 'embed', NS.officeDocRels)));
};

/**
 * Returns the bytes of the image used as this shape's *fill*
 * (`<a:blipFill>` nested under `<p:spPr>`, as written by
 * `setShapeImageFill`). Distinct from `getShapeImageBytes`, which only
 * applies to `<p:pic>` picture shapes.
 *
 * Returns null if the shape has no image fill, the blip has no
 * `r:embed`, or the embed points at an external `r:link`.
 */
export const getShapeImageFillBytes = (shape: SlideShapeData): Uint8Array | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const blipFill = firstChildElement(spPr, qname('a', 'blipFill', NS.dml));
  if (!blipFill) return null;
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) return null;
  const rEmbed = getAttrValue(blip, qname('r', 'embed', NS.officeDocRels));
  if (rEmbed === null) return null;
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;
  const rel = rels.items.find((r) => r.id === rEmbed);
  if (!rel || rel.targetMode === 'External') return null;
  const mediaName = rel.target.startsWith('/')
    ? partName(rel.target)
    : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
  const part = pkg.getPart(mediaName);
  return part?.data ?? null;
};

/**
 * Returns the image format token (`'png'`, `'jpeg'`, …) for whichever
 * image bytes the shape carries — picture (`<p:pic>`) or image-fill
 * (`<a:blipFill>` on `<p:spPr>`). Returns `null` if the shape has no
 * embedded image or the bytes don't match a recognized signature.
 */
export const getShapeImageFormat = (shape: SlideShapeData): ImageFormat | null => {
  const bytes = getShapeImageBytes(shape) ?? getShapeImageFillBytes(shape);
  if (bytes === null) return null;
  return detectImageFormat(bytes);
};

/**
 * Reads the picture's opacity (0–1 fraction). Returns `null` when no
 * `<a:alphaModFix>` is present (PowerPoint treats absence as fully
 * opaque); returns `1` when an explicit alphaModFix sets full opacity.
 */
export const getShapeImageOpacity = (shape: SlideShapeData): number | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') return null;
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) return null;
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) return null;
  const alpha = firstChildElement(blip, qname('a', 'alphaModFix', NS.dml));
  if (!alpha) return null;
  const amt = getAttrValue(alpha, qname('', 'amt', ''));
  if (amt === null) return 1;
  const n = Number.parseInt(amt, 10);
  if (!Number.isFinite(n)) return null;
  return n / 100000;
};

/**
 * Reads the picture's crop fractions. Returns `null` when no
 * `<a:srcRect>` is present; otherwise returns a fully-populated object
 * with every side filled in (0 for omitted sides on disk).
 */
export const getShapeImageCrop = (shape: SlideShapeData): ImageCrop | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') return null;
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) return null;
  const srcRect = firstChildElement(blipFill, qname('a', 'srcRect', NS.dml));
  if (!srcRect) return null;
  const parseSide = (local: string): number => {
    const v = getAttrValue(srcRect, qname('', local, ''));
    if (v === null) return 0;
    const n = Number.parseInt(v, 10);
    return Number.isFinite(n) ? n / 100000 : 0;
  };
  return {
    left: parseSide('l'),
    top: parseSide('t'),
    right: parseSide('r'),
    bottom: parseSide('b'),
  };
};

/**
 * Adjusts the picture's brightness by writing `<a:lumOff val="…"/>`
 * inside `<a:blip>`. The value is a -1..1 fraction:
 *
 *   - `1`     → +100% brightness
 *   - `0` or `null` → no offset (any prior `<a:lumOff>` is removed)
 *   - `-1`    → -100% brightness
 *
 * Throws for non-picture shapes and on values outside [-1, 1].
 *
 * Note: PowerPoint's "Picture Format › Corrections" UI couples this
 * with `<a:lumMod>` for some presets; this primitive sets only
 * `lumOff` to keep the surface honest. Read it back via
 * `getShapeImageBrightness`.
 */
export const setShapeImageBrightness = (shape: SlideShapeData, value: number | null): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') {
    throw new Error(
      `setShapeImageBrightness only works on picture shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) throw new Error('picture has no <p:blipFill>');
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) throw new Error('picture <p:blipFill> has no <a:blip>');
  blip.children = blip.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'lumOff'),
  );
  if (value !== null && value !== 0) {
    if (!Number.isFinite(value) || value < -1 || value > 1) {
      throw new RangeError(`brightness must be in [-1, 1], got ${value}`);
    }
    blip.children.push(
      elem(qname('a', 'lumOff', NS.dml), {
        attrs: [attr(qname('', 'val', ''), String(Math.round(value * 100000)))],
      }),
    );
  }
  commitAndRefresh(shape);
};

/**
 * Adjusts the picture's contrast by writing `<a:lumMod val="…"/>` on
 * `<a:blip>`. The value is a 0..2 fraction:
 *
 *   - `1` or `null` → no modulation (any prior `<a:lumMod>` is removed)
 *   - `0.5`         → 50% of original luminance variance (washed out)
 *   - `1.5`         → 150% (boosted contrast; PowerPoint clamps to
 *                       what the renderer supports)
 *
 * Throws on non-picture shapes and on values outside `[0, 2]`. The
 * primitive maps directly to `ST_PositiveFixedPercentage` × 100000.
 */
export const setShapeImageContrast = (shape: SlideShapeData, value: number | null): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') {
    throw new Error(
      `setShapeImageContrast only works on picture shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) throw new Error('picture has no <p:blipFill>');
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) throw new Error('picture <p:blipFill> has no <a:blip>');
  blip.children = blip.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'lumMod'),
  );
  if (value !== null && value !== 1) {
    if (!Number.isFinite(value) || value < 0 || value > 2) {
      throw new RangeError(`contrast must be in [0, 2], got ${value}`);
    }
    blip.children.push(
      elem(qname('a', 'lumMod', NS.dml), {
        attrs: [attr(qname('', 'val', ''), String(Math.round(value * 100000)))],
      }),
    );
  }
  commitAndRefresh(shape);
};

/**
 * Reads the picture's contrast modulation (the `<a:lumMod>` fraction
 * in [0, 2]). Returns `null` when no `<a:lumMod>` is present.
 */
export const getShapeImageContrast = (shape: SlideShapeData): number | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') return null;
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) return null;
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) return null;
  const lumMod = firstChildElement(blip, qname('a', 'lumMod', NS.dml));
  if (!lumMod) return null;
  const v = getAttrValue(lumMod, qname('', 'val', ''));
  if (v === null) return null;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n / 100000 : null;
};

/**
 * Reads the picture's brightness offset (the `<a:lumOff>` fraction
 * in [-1, 1]). Returns `null` when no `<a:lumOff>` is present.
 */
export const getShapeImageBrightness = (shape: SlideShapeData): number | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') return null;
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) return null;
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) return null;
  const lumOff = firstChildElement(blip, qname('a', 'lumOff', NS.dml));
  if (!lumOff) return null;
  const v = getAttrValue(lumOff, qname('', 'val', ''));
  if (v === null) return null;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n / 100000 : null;
};

export const setShapeImageOpacity = (shape: SlideShapeData, opacity: number | null): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') {
    throw new Error(
      `setShapeImageOpacity only works on picture shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) throw new Error('picture has no <p:blipFill>');
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) throw new Error('picture <p:blipFill> has no <a:blip>');

  blip.children = blip.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        c.name.localName === 'alphaModFix'
      ),
  );

  if (opacity !== null) {
    if (!Number.isFinite(opacity) || opacity < 0 || opacity > 1) {
      throw new RangeError(`opacity must be in [0, 1], got ${opacity}`);
    }
    blip.children.push(
      elem(NAME_ALPHA_MOD_FIX_FN, {
        attrs: [attr(ATTR_AMT_FN, String(Math.round(opacity * 100000)))],
      }),
    );
  }
  commitAndRefresh(shape);
};

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
    (c) => c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'blip',
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
export const setShapeAnimation = (shape: SlideShapeData, opts: AnimationOptions): void => {
  const slide = shape[SHAPE_SLIDE];
  removeExistingTiming(slide);
  const spid = shape[SHAPE_SNAPSHOT].id;
  const timing = buildSingleEffectTiming(spid, opts);
  insertTimingAtEnd(slide, timing);
  commitSlideData(slide);
  refreshSlideData(slide);
};

/**
 * Returns the animation effect bound to this shape via the slide's
 * `<p:timing>` tree, or `null` if the shape has no animation in the
 * v1 single-effect schema we model. Unknown presets are reported as a
 * raw `null` rather than guessing.
 */
export const getShapeAnimation = (shape: SlideShapeData): AnimationEffect | null => {
  const slide = shape[SHAPE_SLIDE];
  const timing = slide[SLIDE_DOCUMENT].root.children.find(
    (c): c is XmlElement =>
      c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'timing',
  );
  if (!timing) return null;

  // Confirm the shape's spid appears in <p:bldLst><p:bldP spid="..."/>.
  const bldLst = firstChildElement(timing, qname('p', 'bldLst', NS.pml));
  if (!bldLst) return null;
  const spidStr = String(shape[SHAPE_SNAPSHOT].id);
  const matched = allChildElements(bldLst, qname('p', 'bldP', NS.pml)).some(
    (b) => getAttrValue(b, qname('', 'spid', '')) === spidStr,
  );
  if (!matched) return null;

  // Walk the timing tree to find the effect cTn for this shape. Our
  // builder emits `<p:cTn presetID="N" presetClass="entr|exit" ...
  // nodeType="clickEffect">` with a `<p:spTgt spid="..."/>` inside. We
  // accept any cTn carrying that combination.
  let presetID: string | null = null;
  let presetClass: string | null = null;
  const walk = (el: XmlElement): boolean => {
    if (el.name.namespaceURI === NS.pml && el.name.localName === 'cTn') {
      const cls = getAttrValue(el, qname('', 'presetClass', ''));
      const id = getAttrValue(el, qname('', 'presetID', ''));
      if (cls && id) {
        // Confirm this cTn targets our shape via a descendant spTgt.
        const targetsShape = (sub: XmlElement): boolean => {
          if (
            sub.name.namespaceURI === NS.pml &&
            sub.name.localName === 'spTgt' &&
            getAttrValue(sub, qname('', 'spid', '')) === spidStr
          ) {
            return true;
          }
          for (const c of sub.children) {
            if (c.kind === 'element' && targetsShape(c)) return true;
          }
          return false;
        };
        if (targetsShape(el)) {
          presetClass = cls;
          presetID = id;
          return true;
        }
      }
    }
    for (const c of el.children) {
      if (c.kind === 'element' && walk(c)) return true;
    }
    return false;
  };
  walk(timing);
  if (!presetID || !presetClass) return null;

  // Map back to AnimationEffect.
  const id = Number.parseInt(presetID, 10);
  if (presetClass === 'entr' && id === 1) return 'appear';
  if (presetClass === 'entr' && id === 10) return 'fadeIn';
  if (presetClass === 'exit' && id === 1) return 'disappear';
  if (presetClass === 'exit' && id === 10) return 'fadeOut';
  return null;
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
 * Returns the slide's body text, or `null` if no `body` placeholder
 * is present. Mirror of `getSlideTitle`; pairs with `setSlideBody`.
 */
export const getSlideBody = (slide: SlideData): string | null => {
  const bodyShape = findSlidePlaceholder(slide, 'body');
  if (bodyShape === null) return null;
  return bodyShape[SHAPE_SNAPSHOT].text ?? null;
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

/**
 * Bulk-fills slide placeholders by type token. Each entry in
 * `byType` maps a `<p:ph type>` token (e.g. `'title'`, `'body'`,
 * `'ftr'`, `'dt'`) to the text to set. Silently skips entries
 * whose placeholder isn't present on the slide.
 *
 * Useful for template-fill workflows where the caller has all the
 * data in one struct.
 */
export const setSlidePlaceholders = (
  slide: SlideData,
  byType: Readonly<Record<string, string>>,
): void => {
  for (const [type, text] of Object.entries(byType)) {
    const shape =
      type === 'title'
        ? (findSlidePlaceholder(slide, 'title') ?? findSlidePlaceholder(slide, 'ctrTitle'))
        : findSlidePlaceholder(slide, type);
    if (shape !== null) setShapeText(shape, text);
  }
};

/**
 * Writes `text` into the first body placeholder on the slide.
 * Newlines start a new paragraph (each becomes its own bullet on
 * layouts that bullet their body placeholder).
 *
 * Throws when the slide has no body placeholder — pair with
 * `findSlideLayoutByType(pres, 'obj')` / `'tx'` to add the slide
 * onto a layout that has one.
 */
export const setSlideBody = (slide: SlideData, text: string): void => {
  const bodyShape = findSlidePlaceholder(slide, 'body');
  if (bodyShape === null) {
    throw new Error('setSlideBody: slide has no body placeholder');
  }
  setShapeText(bodyShape, text);
};
