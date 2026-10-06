// Slide comments.
import { getSlides } from './slide-query.ts';

import {
  type PartName,
  emptyRels,
  nextRelId,
  partName,
  resolveTarget,
} from '../../internal/opc/index.ts';
import type { OpcPackage, Part } from '../../internal/parts/index.ts';
import {
  DEFAULT_COMMENT_POSITION,
  CREATION_ID_URI,
  MODERN_AUTHORS_CONTENT_TYPE,
  MODERN_COMMENTS_CONTENT_TYPE,
  P14_NS,
  REL_TYPES,
  type CommentAuthor,
  type CommentPosition,
  type CommentStatus,
  type ModernAuthor,
  type ModernComment,
  type SlideComment,
  anchorOf,
  appendModernReply,
  buildCommentAuthorListDoc,
  buildCommentListDoc,
  buildModernAuthorElement,
  buildModernAuthorListDoc,
  buildModernCommentElement,
  buildModernCommentListDoc,
  buildModernReplyElement,
  buildSlideAnchor,
  buildUnknownAnchor,
  findModernComment,
  findModernReply,
  modernCommentElements,
  readCommentAuthorList,
  readCommentList,
  readModernAuthorList,
  readModernCommentList,
  removeModernComment,
  setModernStatus,
  setModernText,
} from '../../internal/presentationml/index.ts';
import {
  NS,
  allChildElements,
  cloneElement,
  firstChildElement,
  getAttrValue,
  qname,
  qnameEquals,
  parseXml,
  serializeXml,
  type XmlDocument,
  type XmlElement,
} from '../../internal/xml/index.ts';
import { newGuid } from '../../internal/bounds.ts';
import {
  COMMENT_MODERN,
  COMMENT_PARENT,
  COMMENT_SLIDE,
  COMMENT_SNAPSHOT,
  INTERNAL_PACKAGE,
  type PresentationData,
  SLIDE_DOCUMENT,
  SLIDE_PART_NAME,
  type SlideCommentData,
  type SlideData,
} from '../_internal-symbols.ts';
import { PRES_PART_NAME, decode, encode } from './_helpers.ts';
// ---------------------------------------------------------------------------
// Comments.
//
// Legacy schema (ECMA-376 Part 1 §19.4):
//   * One package-level `/ppt/commentAuthors.xml` holds every author.
//   * One `/ppt/comments/comment{N}.xml` per slide that has comments;
//     Newly created parts use the slide number when available. Imported
//     part names are resolved through relationships.
//   * Slide rels reference the slide's comments part; presentation rels
//     reference the author list.
//
// Authors are deduped by (name, initials). `idx` allocation is per-author
// monotonic; we read each author's `lastIdx` and bump it on add.

const COMMENT_AUTHORS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml';
const COMMENTS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.comments+xml';

const relatedPart = (pkg: OpcPackage, source: PartName, type: string): PartName | null => {
  const rel = pkg
    .getRels(source)
    ?.items.find((r) => r.type === type && r.targetMode !== 'External');
  return rel ? resolveTarget(source, rel.target) : null;
};

const commentsPartNameForSlide = (slide: SlideData): PartName | null =>
  relatedPart(slide[INTERNAL_PACKAGE], slide[SLIDE_PART_NAME], REL_TYPES.comments);

const unusedPartName = (pkg: OpcPackage, base: string): PartName => {
  let candidate = partName(`${base}.xml`);
  let suffix = 1;
  while (pkg.getPart(candidate)) candidate = partName(`${base}-${suffix++}.xml`);
  return candidate;
};

const loadAuthorList = (pkg: OpcPackage): CommentAuthor[] => {
  const name = relatedPart(pkg, PRES_PART_NAME, REL_TYPES.commentAuthors);
  const part = name ? pkg.getPart(name) : null;
  if (part === null) return [];
  const list = readCommentAuthorList(parseXml(decode(part.data)).root);
  return list.authors.slice();
};

// Retain imported nodes and extension XML while adding/removing keyed entries.
const reconcileList = (
  original: XmlDocument,
  generated: XmlDocument,
  localName: string,
  key: (node: XmlElement) => string,
  updateAttributes: boolean,
): XmlDocument => {
  const pending = new Map(
    generated.root.children
      .filter((node): node is XmlElement => node.kind === 'element')
      .map((node) => [key(node), node]),
  );
  original.root.children = original.root.children.filter((node) => {
    if (
      node.kind !== 'element' ||
      node.name.namespaceURI !== NS.pml ||
      node.name.localName !== localName
    )
      return true;
    const replacement = pending.get(key(node));
    if (!replacement) return false;
    pending.delete(key(node));
    if (updateAttributes) {
      for (const attribute of replacement.attrs) {
        const index = node.attrs.findIndex((existing) =>
          qnameEquals(existing.name, attribute.name),
        );
        if (index < 0) node.attrs.push(attribute);
        else node.attrs[index] = { ...node.attrs[index]!, value: attribute.value };
      }
    }
    return true;
  });
  const added = [...pending.values()];
  for (const node of added) node.prefixDecls.set('p', NS.pml);
  const extension = original.root.children.findIndex(
    (node) =>
      node.kind === 'element' &&
      node.name.namespaceURI === NS.pml &&
      node.name.localName === 'extLst',
  );
  original.root.children.splice(
    extension < 0 ? original.root.children.length : extension,
    0,
    ...added,
  );
  return original;
};

const writeAuthorList = (pkg: OpcPackage, authors: ReadonlyArray<CommentAuthor>): void => {
  const doc = buildCommentAuthorListDoc(authors);
  const bytes = encode(serializeXml(doc));
  const name =
    relatedPart(pkg, PRES_PART_NAME, REL_TYPES.commentAuthors) ??
    unusedPartName(pkg, '/ppt/commentAuthors');
  const existing = pkg.getPart(name);
  if (existing !== null) {
    const preserved = reconcileList(
      parseXml(decode(existing.data)),
      doc,
      'cmAuthor',
      (node) => getAttrValue(node, qname('', 'id', '')) ?? '',
      true,
    );
    existing.data = encode(serializeXml(preserved));
    return;
  }
  pkg.addPart(name, COMMENT_AUTHORS_CONTENT_TYPE, bytes);
  // presentation → commentAuthors rel.
  const presRels = pkg.getRels(PRES_PART_NAME) ?? emptyRels();
  const exists = presRels.items.some(
    (r) => r.type === REL_TYPES.commentAuthors && r.targetMode !== 'External',
  );
  if (!exists) {
    presRels.items.push({
      id: nextRelId(presRels.items.map((r) => r.id)),
      type: REL_TYPES.commentAuthors,
      target: name,
      targetMode: 'Internal',
    });
    pkg.setRels(PRES_PART_NAME, presRels);
  }
};

const loadCommentsForSlide = (slide: SlideData): SlideComment[] => {
  const pkg = slide[INTERNAL_PACKAGE];
  const partNameValue = commentsPartNameForSlide(slide);
  const part = partNameValue ? pkg.getPart(partNameValue) : null;
  if (part === null) return [];
  const list = readCommentList(parseXml(decode(part.data)).root);
  return list.comments.slice();
};

const writeCommentsForSlide = (slide: SlideData, comments: ReadonlyArray<SlideComment>): void => {
  const pkg = slide[INTERNAL_PACKAGE];
  const relatedName = commentsPartNameForSlide(slide);
  const slideN = slide[SLIDE_PART_NAME].match(/slide(\d+)\.xml$/)?.[1] ?? '1';
  const commentsName = relatedName ?? unusedPartName(pkg, `/ppt/comments/comment${slideN}`);

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
    const preserved = reconcileList(
      parseXml(decode(existing.data)),
      doc,
      'cm',
      (node) =>
        `${getAttrValue(node, qname('', 'authorId', ''))}:${getAttrValue(node, qname('', 'idx', ''))}`,
      false,
    );
    existing.data = encode(serializeXml(preserved));
    return;
  }
  pkg.addPart(commentsName, COMMENTS_CONTENT_TYPE, bytes);

  const slideRels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const hasRel = slideRels.items.some(
    (r) => r.type === REL_TYPES.comments && r.targetMode !== 'External',
  );
  if (!hasRel) {
    slideRels.items.push({
      id: nextRelId(slideRels.items.map((r) => r.id)),
      type: REL_TYPES.comments,
      target: commentsName,
      targetMode: 'Internal',
    });
    pkg.setRels(slide[SLIDE_PART_NAME], slideRels);
  }
};

const commentKey = (comment: { authorId: number; idx: number }): string =>
  `${comment.authorId}:${comment.idx}`;

const asCommentData = (
  slide: SlideData,
  snap: SlideComment,
  author: CommentAuthor,
): SlideCommentData => ({
  [COMMENT_PARENT]: null,
  [COMMENT_SLIDE]: slide,
  [COMMENT_SNAPSHOT]: snap,
  [COMMENT_MODERN]: null,
  author,
});

// ---------------------------------------------------------------------------
// Modern comments ([MS-PPTX] §2.16.1).
//
// A second, incompatible way to say the same thing: PowerPoint 2021 and
// Microsoft 365 write `<p188:cm>` threads into a per-slide part related from
// the slide, with the authors in one `/ppt/authors.xml`. The differences that
// matter here are that a thread owns its replies rather than the replies
// pointing back, and that a thread can be resolved.
//
// Both formats read through `getSlideComments`; every mutation checks which
// one a handle came from first. A handle is never converted between them —
// a reply goes where its thread is, and a new thread goes where the slide
// already keeps its comments.

/** Where a brand-new modern comment part goes when a slide has none. */
const modernPartBase = (slide: SlideData): string => {
  const n = slide[SLIDE_PART_NAME].match(/slide(\d+)\.xml$/)?.[1] ?? '1';
  return `/ppt/comments/modernComment${n}`;
};

const modernAuthorsName = (pkg: OpcPackage): PartName | null =>
  relatedPart(pkg, PRES_PART_NAME, REL_TYPES.authors);

const modernCommentsName = (slide: SlideData): PartName | null =>
  relatedPart(slide[INTERNAL_PACKAGE], slide[SLIDE_PART_NAME], REL_TYPES.modernComment);

const loadModernAuthors = (pkg: OpcPackage): ReadonlyArray<ModernAuthor> => {
  const name = modernAuthorsName(pkg);
  const part = name ? pkg.getPart(name) : null;
  if (part === null) return [];
  return readModernAuthorList(parseXml(decode(part.data)).root);
};

/** The parsed modern comment part for a slide, kept open for mutation. */
const openModernComments = (slide: SlideData): { part: Part; doc: XmlDocument } | null => {
  const name = modernCommentsName(slide);
  const part = name ? slide[INTERNAL_PACKAGE].getPart(name) : null;
  if (part === null) return null;
  return { part, doc: parseXml(decode(part.data)) };
};

const saveModernComments = (slide: SlideData, open: { part: Part; doc: XmlDocument }): void => {
  // A part with no threads left is an orphan: PowerPoint writes neither it
  // nor the relationship that reaches it.
  if (modernCommentElements(open.doc.root).length === 0) {
    const pkg = slide[INTERNAL_PACKAGE];
    const name = modernCommentsName(slide);
    if (name !== null) pkg.removePart(name);
    const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
    if (rels !== null) {
      const before = rels.items.length;
      rels.items = rels.items.filter((r) => r.type !== REL_TYPES.modernComment);
      if (rels.items.length !== before) pkg.setRels(slide[SLIDE_PART_NAME], rels);
    }
    return;
  }
  open.part.data = encode(serializeXml(open.doc));
};

/** Adds one author to `/ppt/authors.xml`, creating the part when it is the first. */
const writeModernAuthor = (pkg: OpcPackage, author: ModernAuthor): void => {
  const name = modernAuthorsName(pkg);
  const part = name ? pkg.getPart(name) : null;
  if (part !== null) {
    const doc = parseXml(decode(part.data));
    doc.root.children.push(buildModernAuthorElement(author));
    part.data = encode(serializeXml(doc));
    return;
  }
  const created = unusedPartName(pkg, '/ppt/authors');
  pkg.addPart(
    created,
    MODERN_AUTHORS_CONTENT_TYPE,
    encode(serializeXml(buildModernAuthorListDoc([author]))),
  );
  const rels = pkg.getRels(PRES_PART_NAME) ?? emptyRels();
  rels.items.push({
    id: nextRelId(rels.items.map((r) => r.id)),
    type: REL_TYPES.authors,
    target: created,
    targetMode: 'Internal',
  });
  pkg.setRels(PRES_PART_NAME, rels);
};

/** The modern author with this name, or a newly registered one. */
const modernAuthorFor = (
  pkg: OpcPackage,
  who: { name: string; initials?: string },
): ModernAuthor => {
  const initials = who.initials ?? (who.name.length > 0 ? who.name.charAt(0) : '?');
  const existing = loadModernAuthors(pkg).find(
    (author) => author.name === who.name && (author.initials ?? '') === initials,
  );
  if (existing) return existing;
  // `userId` and `providerId` are required by the schema and name an account
  // in an identity provider. This library has neither, and inventing one
  // would attribute the comment to somebody. Empty strings say as much.
  const author: ModernAuthor = {
    id: newGuid(),
    name: who.name,
    initials,
    userId: '',
    providerId: '',
  };
  writeModernAuthor(pkg, author);
  return author;
};

/** The slide's `<p14:creationId>`, which an anchor names the slide by. */
const slideCreationId = (slide: SlideData): number | null => {
  const extLst = firstChildElement(slide[SLIDE_DOCUMENT].root, qname('p', 'extLst', NS.pml));
  if (extLst === null) return null;
  for (const ext of allChildElements(extLst, qname('p', 'ext', NS.pml))) {
    if (getAttrValue(ext, qname('', 'uri', '')) !== CREATION_ID_URI) continue;
    const creationId = firstChildElement(ext, qname('p14', 'creationId', P14_NS));
    const value = creationId && getAttrValue(creationId, qname('', 'val', ''));
    const parsed = value === null || value === undefined ? Number.NaN : Number.parseInt(value, 10);
    if (Number.isFinite(parsed)) return parsed;
  }
  return null;
};

/** The slide's id as `<p:sldIdLst>` gives it. */
const slideIdOf = (slide: SlideData): number | null => {
  const pkg = slide[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  const rels = pkg.getRels(PRES_PART_NAME);
  if (presPart === null || rels === null) return null;
  const rel = rels.items.find(
    (item) =>
      item.type === REL_TYPES.slide &&
      item.targetMode !== 'External' &&
      resolveTarget(PRES_PART_NAME, item.target) === slide[SLIDE_PART_NAME],
  );
  if (!rel) return null;
  const sldIdLst = firstChildElement(
    parseXml(decode(presPart.data)).root,
    qname('p', 'sldIdLst', NS.pml),
  );
  if (sldIdLst === null) return null;
  for (const sldId of allChildElements(sldIdLst, qname('p', 'sldId', NS.pml))) {
    if (getAttrValue(sldId, qname('r', 'id', NS.officeDocRels)) !== rel.id) continue;
    const parsed = Number.parseInt(getAttrValue(sldId, qname('', 'id', '')) ?? '', 10);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
};

/**
 * The anchor for a new thread. An existing thread on the same slide already
 * names it exactly, so that one is copied; otherwise the moniker is built
 * from the slide's own identifiers. A slide that carries no creation id —
 * every PowerPoint slide does, decks built from scratch do not — gets the
 * schema's own `unknownAnchor` rather than a made-up moniker.
 */
const anchorForSlide = (slide: SlideData, root: XmlElement): XmlElement => {
  for (const thread of modernCommentElements(root)) {
    const existing = anchorOf(thread);
    if (existing !== null) return cloneElement(existing);
  }
  const creationId = slideCreationId(slide);
  const slideId = slideIdOf(slide);
  if (creationId === null || slideId === null) return buildUnknownAnchor();
  return buildSlideAnchor(creationId, slideId);
};

/** Opens the slide's modern comment part, creating an empty one if needed. */
const openOrCreateModernComments = (slide: SlideData): { part: Part; doc: XmlDocument } => {
  const existing = openModernComments(slide);
  if (existing !== null) return existing;
  const pkg = slide[INTERNAL_PACKAGE];
  const name = unusedPartName(pkg, modernPartBase(slide));
  const doc = buildModernCommentListDoc([]);
  pkg.addPart(name, MODERN_COMMENTS_CONTENT_TYPE, encode(serializeXml(doc)));
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  rels.items.push({
    id: nextRelId(rels.items.map((r) => r.id)),
    type: REL_TYPES.modernComment,
    target: name,
    targetMode: 'Internal',
  });
  pkg.setRels(slide[SLIDE_PART_NAME], rels);
  const part = pkg.getPart(name);
  if (part === null) throw new Error('addSlideComment: the new comment part went missing');
  return { part, doc };
};

/**
 * Every author known to the package's `commentAuthors.xml`.
 * Returns an empty array when no author list exists.
 */
export const getCommentAuthors = (pres: PresentationData): ReadonlyArray<CommentAuthor> =>
  loadAuthorList(pres[INTERNAL_PACKAGE]);

/**
 * Returns every slide that has at least one comment by the given
 * author name. Accepts a literal string (exact-equality) or a
 * `RegExp` for pattern matches. Sibling of `findCommentsByAuthor`
 * (which returns the comments themselves).
 */
export const findSlidesWithCommentsByAuthor = (
  pres: PresentationData,
  authorName: string | RegExp,
): ReadonlyArray<SlideData> => {
  const matches =
    typeof authorName === 'string'
      ? (n: string) => n === authorName
      : (n: string) => authorName.test(n);
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (getSlideComments(slide).some((c) => matches(c.author.name))) out.push(slide);
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
 * Histogram of comment counts by author display name across the whole
 * deck. Useful for "who reviewed this deck the most?" audits.
 * Authors with the same display name (a real-world case for shared
 * mailbox identities) get merged into the same bucket; pair with
 * `getPresentationCommenters` when you need to keep authors with
 * identical names separate by `id`.
 */
export const getPresentationCommentCountsByAuthor = (
  pres: PresentationData,
): Readonly<Record<string, number>> => {
  const counts: Record<string, number> = {};
  for (const slide of getSlides(pres)) {
    for (const c of getSlideComments(slide)) {
      const name = c.author.name;
      counts[name] = (counts[name] ?? 0) + 1;
    }
  }
  return counts;
};

/**
 * Dense histogram of comment counts by 0-based slide index. Every
 * slide in the deck appears as an element (count `0` when the slide
 * has no comments), so the array shape is dense — handy for charting
 * comment density per slide without re-indexing.
 */
export const getPresentationCommentCountsBySlide = (
  pres: PresentationData,
): ReadonlyArray<number> => {
  const slides = getSlides(pres);
  return slides.map((s) => getSlideComments(s).length);
};

/**
 * Looks up a `CommentAuthor` from `commentAuthors.xml` by display
 * name. Accepts a literal string (exact equality) or a `RegExp` for
 * pattern matches. Returns `null` when no author matches. Sibling of
 * `findCommentsByAuthor` — the latter returns the matching comments;
 * this returns the author handle for downstream metadata reads (id,
 * initials, color).
 */
export const findCommentAuthorByName = (
  pres: PresentationData,
  authorName: string | RegExp,
): CommentAuthor | null => {
  for (const a of getCommentAuthors(pres)) {
    const hit = typeof authorName === 'string' ? a.name === authorName : authorName.test(a.name);
    if (hit) return a;
  }
  return null;
};

/**
 * Returns every comment whose author name matches `authorName`
 * across every slide in the deck. Accepts a literal string (exact-
 * equality) or a `RegExp` for pattern matches — useful for reviewer-
 * specific filters ("show me all of Alice's notes" / `/^review-bot/`).
 */
export const findCommentsByAuthor = (
  pres: PresentationData,
  authorName: string | RegExp,
): ReadonlyArray<SlideCommentData> => {
  const out: SlideCommentData[] = [];
  for (const slide of getSlides(pres)) {
    for (const c of getSlideComments(slide)) {
      const hit =
        typeof authorName === 'string'
          ? c.author.name === authorName
          : authorName.test(c.author.name);
      if (hit) out.push(c);
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
 * Reads the slide's modern threads as comment handles: the thread first,
 * then its replies, so a reply always follows the comment it answers.
 *
 * `authorId` and `idx` on the snapshot are the author's place in the modern
 * author list and a running count — see `SlideCommentData`. The real
 * identifiers are GUIDs and live on `COMMENT_MODERN`.
 */
const readModernSlideComments = (slide: SlideData): SlideCommentData[] => {
  const open = openModernComments(slide);
  if (open === null) return [];
  const authors = loadModernAuthors(slide[INTERNAL_PACKAGE]);
  const authorAt = new Map(authors.map((author, index) => [author.id, index]));
  const out: SlideCommentData[] = [];
  let counter = 0;
  const handleOf = (
    thread: ModernComment,
    snapshot: ModernComment | ModernComment['replies'][number],
    position: CommentPosition | null,
    parent: SlideCommentData | null,
  ): SlideCommentData => {
    const index = authorAt.get(snapshot.authorId) ?? -1;
    const author = authors[index];
    counter += 1;
    const data: SlideCommentData = {
      [COMMENT_PARENT]: parent,
      [COMMENT_SLIDE]: slide,
      [COMMENT_SNAPSHOT]: {
        ...(parent ? { parent: { authorId: -1, idx: -1 } } : {}),
        authorId: index,
        idx: counter,
        dt: snapshot.created,
        text: snapshot.text,
        position,
      },
      [COMMENT_MODERN]: {
        threadId: thread.id,
        id: snapshot.id,
        status: snapshot.status,
        authorId: snapshot.authorId,
      },
      // An author the list does not name is surfaced as an empty record
      // rather than dropping the comment, the same way the legacy path does.
      author: {
        id: index,
        name: author?.name ?? '',
        initials: author?.initials ?? '',
        lastIdx: 0,
        clrIdx: null,
      },
    };
    return data;
  };
  for (const thread of readModernCommentList(open.doc.root)) {
    const root = handleOf(thread, thread, thread.position, null);
    out.push(root);
    for (const reply of thread.replies) out.push(handleOf(thread, reply, null, root));
  }
  return out;
};

/**
 * Returns every comment attached to the slide, with the author already
 * resolved. Both formats are read: the ECMA-376 `<p:cm>` list first, then
 * the modern threads. The list is read-only — use `addSlideComment` /
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
  const byId = new Map(out.map((comment) => [commentKey(comment[COMMENT_SNAPSHOT]), comment]));
  for (const comment of out) {
    const parent = comment[COMMENT_SNAPSHOT].parent;
    comment[COMMENT_PARENT] = parent ? (byId.get(commentKey(parent)) ?? null) : null;
  }
  return [...out, ...readModernSlideComments(slide)];
};

/** Options shared by both comment formats. */
interface AddCommentOptions {
  author: { name: string; initials?: string };
  text: string;
  position?: CommentPosition | null;
  date?: Date;
  /** Existing comment on this slide to reply to. */
  replyTo?: SlideCommentData;
}

/** Writes a modern thread, or a reply inside an existing one. */
const addModernComment = (slide: SlideData, opts: AddCommentOptions): SlideCommentData => {
  const pkg = slide[INTERNAL_PACKAGE];
  const created = (opts.date ?? new Date()).toISOString();
  const author = modernAuthorFor(pkg, opts.author);
  const open = openOrCreateModernComments(slide);
  const id = newGuid();
  const parent = opts.replyTo?.[COMMENT_MODERN] ?? null;
  if (parent !== null) {
    const thread = findModernComment(open.doc.root, parent.threadId);
    if (thread === null) throw new Error('addSlideComment: reply parent must exist on this slide');
    appendModernReply(
      thread,
      buildModernReplyElement({ id, authorId: author.id, created, text: opts.text }),
    );
  } else {
    open.doc.root.children.push(
      buildModernCommentElement({
        id,
        authorId: author.id,
        created,
        text: opts.text,
        anchor: anchorForSlide(slide, open.doc.root),
        ...(opts.position ? { position: opts.position } : {}),
      }),
    );
  }
  saveModernComments(slide, open);
  // Re-read so the handle carries the same projection every other reader
  // sees, rather than one built twice from two places.
  const added = readModernSlideComments(slide).find(
    (comment) => comment[COMMENT_MODERN]?.id === id,
  );
  if (!added) throw new Error('addSlideComment: the new comment did not come back');
  return added;
};

/**
 * Adds a comment to the slide. Returns the new comment handle.
 *
 * Which format it is written in follows the file rather than the caller: a
 * reply joins its own thread, and a new thread goes where the slide — or,
 * failing that, the deck — already keeps its comments. Only a deck with no
 * comments at all starts an ECMA-376 `<p:cm>` list, which everything can
 * read.
 *
 * Author handling: if an author with the given `name`+`initials` already
 * exists in the relevant author list, the existing record is reused (and, in
 * the legacy list, its `lastIdx` is bumped). Otherwise a new author is
 * allocated. `initials` defaults to the first character of `name`.
 *
 * `position` is the comment pin, in EMU. The legacy schema makes `<p:pos>`
 * mandatory, so a legacy comment added without one is pinned to the slide's
 * origin rather than written without the element; the modern schema makes it
 * optional, so a modern comment added without one has no pin.
 *
 * `date` defaults to the current time.
 */
export const addSlideComment = (slide: SlideData, opts: AddCommentOptions): SlideCommentData => {
  if (opts.replyTo && opts.replyTo[COMMENT_SLIDE] !== slide) {
    throw new Error('addSlideComment: reply parent must exist on this slide');
  }
  const usesModern =
    opts.replyTo !== undefined
      ? opts.replyTo[COMMENT_MODERN] !== null
      : modernCommentsName(slide) !== null ||
        (commentsPartNameForSlide(slide) === null &&
          modernAuthorsName(slide[INTERNAL_PACKAGE]) !== null);
  if (usesModern) return addModernComment(slide, opts);
  const comments = loadCommentsForSlide(slide);
  const parent = opts.replyTo?.[COMMENT_SNAPSHOT];
  if (opts.replyTo && !comments.some((comment) => commentKey(comment) === commentKey(parent!))) {
    throw new Error('addSlideComment: reply parent must exist on this slide');
  }
  const dt = (opts.date ?? new Date()).toISOString();
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

  const snap: SlideComment = {
    ...(parent ? { parent: { authorId: parent.authorId, idx: parent.idx } } : {}),
    authorId: updatedAuthor.id,
    idx: newIdx,
    dt,
    text: opts.text,
    position: opts.position ?? DEFAULT_COMMENT_POSITION,
  };

  comments.push(snap);
  writeCommentsForSlide(slide, comments);

  const handle = asCommentData(slide, snap, updatedAuthor);
  handle[COMMENT_PARENT] = opts.replyTo ?? null;
  return handle;
};

/**
 * The element a modern handle points at, together with its open part.
 * Throws where the legacy path throws: a handle whose comment is no longer
 * in the file is a caller error, not something to paper over.
 */
const openModernTarget = (
  comment: SlideCommentData,
  caller: string,
): { open: { part: Part; doc: XmlDocument }; element: XmlElement } => {
  const reference = comment[COMMENT_MODERN];
  if (reference === null) throw new Error(`${caller}: not a modern comment`);
  const open = openModernComments(comment[COMMENT_SLIDE]);
  const thread = open === null ? null : findModernComment(open.doc.root, reference.threadId);
  if (open === null || thread === null) throw new Error(`${caller}: comment no longer exists`);
  const element =
    reference.id === reference.threadId ? thread : findModernReply(thread, reference.id);
  if (element === null) throw new Error(`${caller}: comment no longer exists`);
  return { open, element };
};

/** Updates comment text while retaining author, date, position, and extension XML. */
export const setCommentText = (comment: SlideCommentData, text: string): void => {
  if (comment[COMMENT_MODERN] !== null) {
    const { open, element } = openModernTarget(comment, 'setCommentText');
    setModernText(element, text);
    saveModernComments(comment[COMMENT_SLIDE], open);
    comment[COMMENT_SNAPSHOT] = { ...comment[COMMENT_SNAPSHOT], text };
    return;
  }
  const slide = comment[COMMENT_SLIDE];
  const snapshot = comment[COMMENT_SNAPSHOT];
  const name = commentsPartNameForSlide(slide);
  const part = name ? slide[INTERNAL_PACKAGE].getPart(name) : null;
  if (!part) throw new Error('setCommentText: comment no longer exists');
  const doc = parseXml(decode(part.data));
  const element = doc.root.children.find(
    (node) =>
      node.kind === 'element' &&
      node.name.namespaceURI === NS.pml &&
      node.name.localName === 'cm' &&
      getAttrValue(node, qname('', 'authorId', '')) === String(snapshot.authorId) &&
      getAttrValue(node, qname('', 'idx', '')) === String(snapshot.idx),
  );
  if (!element || element.kind !== 'element')
    throw new Error('setCommentText: comment no longer exists');
  const body = element.children.find(
    (node) =>
      node.kind === 'element' &&
      node.name.namespaceURI === NS.pml &&
      node.name.localName === 'text',
  );
  if (!body || body.kind !== 'element')
    throw new Error('setCommentText: comment text element is missing');
  body.children = [{ kind: 'text', data: text }];
  part.data = encode(serializeXml(doc));
  comment[COMMENT_SNAPSHOT] = { ...snapshot, text };
};

/**
 * Removes the comment and all its replies from its slide's comments part. If the comment
 * was the last one on the slide, the comments part and the
 * slide → comments rel are also removed. The author entry in
 * `commentAuthors.xml` is left intact (an author may have comments on
 * other slides).
 */
export const removeSlideComment = (comment: SlideCommentData): void => {
  const reference = comment[COMMENT_MODERN];
  if (reference !== null) {
    const open = openModernComments(comment[COMMENT_SLIDE]);
    if (open === null) throw new Error('removeSlideComment: comment no longer exists');
    // Removing a thread takes its replies with it, which is what the element
    // holding them already means.
    if (!removeModernComment(open.doc.root, reference.id)) {
      throw new Error('removeSlideComment: comment no longer exists');
    }
    saveModernComments(comment[COMMENT_SLIDE], open);
    return;
  }
  const slide = comment[COMMENT_SLIDE];
  const target = comment[COMMENT_SNAPSHOT];
  const comments = loadCommentsForSlide(slide);
  const children = new Map<string, string[]>();
  for (const item of comments)
    if (item.parent) {
      const key = commentKey(item.parent);
      const siblings = children.get(key) ?? [];
      siblings.push(commentKey(item));
      children.set(key, siblings);
    }
  const removed = new Set<string>();
  const queue = [commentKey(target)];
  while (queue.length) {
    const key = queue.pop()!;
    if (removed.has(key)) continue;
    removed.add(key);
    for (const child of children.get(key) ?? []) queue.push(child);
  }
  const remaining = comments.filter((item) => !removed.has(commentKey(item)));
  writeCommentsForSlide(slide, remaining);
};

/**
 * Strips every comment from every slide in the deck. Returns the
 * number of comments removed. Built on `writeCommentsForSlide`,
 * so each slide's legacy comment part is rewritten with an empty
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
  if (comments.length > 0) writeCommentsForSlide(slide, []);
  const modern = readModernSlideComments(slide);
  if (modern.length > 0) {
    const open = openModernComments(slide);
    if (open !== null) {
      open.doc.root.children = [];
      saveModernComments(slide, open);
    }
  }
  return comments.length + modern.length;
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

/** Parent handle from the same read, or null for roots and missing parents. */
export const getCommentParent = (comment: SlideCommentData): SlideCommentData | null =>
  comment[COMMENT_PARENT];

/**
 * Which of the two comment formats this handle came from. `'legacy'` is the
 * ECMA-376 `<p:cm>` list every reader understands; `'modern'` is the
 * `<p188:cm>` thread PowerPoint 2021 and Microsoft 365 write.
 */
export const getCommentFormat = (comment: SlideCommentData): 'legacy' | 'modern' =>
  comment[COMMENT_MODERN] === null ? 'legacy' : 'modern';

/**
 * Whether the thread has been resolved. `null` where the question does not
 * apply or cannot be answered: a legacy comment, which has nowhere to record
 * it, or a modern one whose file states a status this schema version does not
 * define.
 */
export const getCommentStatus = (comment: SlideCommentData): CommentStatus | null =>
  comment[COMMENT_MODERN]?.status ?? null;

/**
 * Resolves a thread, reopens it (`'active'`), or closes it. Replies carry a
 * status of their own, so this works on any modern comment.
 *
 * Throws for a legacy comment: the ECMA-376 schema has no place to put it,
 * and quietly doing nothing would leave the caller believing a thread had
 * been resolved.
 */
export const setCommentStatus = (comment: SlideCommentData, status: CommentStatus): void => {
  if (comment[COMMENT_MODERN] === null) {
    throw new Error(
      'setCommentStatus: this comment is in the ECMA-376 format, which cannot record a status',
    );
  }
  const { open, element } = openModernTarget(comment, 'setCommentStatus');
  setModernStatus(element, status);
  saveModernComments(comment[COMMENT_SLIDE], open);
  comment[COMMENT_MODERN] = { ...comment[COMMENT_MODERN], status };
};
