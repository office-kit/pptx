// Modern comments — `/ppt/authors.xml` and one comment part per slide.
//
// This is Microsoft's own format, not ECMA-376: [MS-PPTX] §2.16.1 (Author
// Part and Comment Part) with the schema in §5.14, target namespace
// `http://schemas.microsoft.com/office/powerpoint/2018/8/main`. PowerPoint
// 2021 and Microsoft 365 write comments this way; everything older writes
// the ECMA `<p:cm>` list in `comments-part.ts`, and one package can carry
// both at once.
//
//   <p188:authorLst>
//     <p188:author id="{GUID}" name="Reviewer" initials="R"
//                  userId="" providerId=""/>
//   </p188:authorLst>
//
//   <p188:cmLst>
//     <p188:cm id="{GUID}" authorId="{GUID}" created="2026-05-15T12:00:00Z"
//              status="resolved">
//       <pc:sldMkLst><pc:docMk/><pc:sldMk cId="1" sldId="256"/></pc:sldMkLst>
//       <p188:replyLst>
//         <p188:reply id="{GUID}" authorId="{GUID}" created="…">…</p188:reply>
//       </p188:replyLst>
//       <p188:txBody><a:bodyPr/><a:p><a:r><a:t>…</a:t></a:r></a:p></p188:txBody>
//     </p188:cm>
//   </p188:cmLst>
//
// Two things the older format does not have:
//
//   * a thread is one `<p188:cm>` holding its replies, rather than separate
//     comments pointing at a parent through the p15 extension;
//   * `status` says whether the thread has been resolved.
//
// The schema declares `pos`, `replyLst`, `txBody` and `extLst` locally, so
// they sit in the p188 namespace even where their *types* come from
// DrawingML. Their order inside `<p188:cm>` is a sequence and is fixed:
// anchor, `pos`, `replyLst`, `txBody`, `extLst`.

import {
  NS,
  type XmlDocument,
  type XmlElement,
  allChildElements,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  qname,
} from '../xml/index.ts';
import { textBodyText } from '../drawingml/text-body.ts';
import { setTextBody } from '../drawingml/text-body-mutation.ts';

/** [MS-PPTX] §5.14 — the modern comment and author namespace. */
export const MODERN_COMMENTS_NS = 'http://schemas.microsoft.com/office/powerpoint/2018/8/main';
/** The content-moniker namespace an anchor is written in ([MS-PPTX] §5.12). */
export const PC_NS = 'http://schemas.microsoft.com/office/powerpoint/2013/main/command';
/** Slide creation ids live in the 2010 namespace ([MS-PPTX] `creationId`). */
export const P14_NS = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
/** `<p:ext>` uri that carries `<p14:creationId>` on a slide. */
export const CREATION_ID_URI = '{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}';

export const MODERN_COMMENTS_CONTENT_TYPE = 'application/vnd.ms-powerpoint.comments+xml';
export const MODERN_AUTHORS_CONTENT_TYPE = 'application/vnd.ms-powerpoint.authors+xml';

// QNames -------------------------------------------------------------------

const NAME_AUTHOR_LST = qname('p188', 'authorLst', MODERN_COMMENTS_NS);
const NAME_AUTHOR = qname('p188', 'author', MODERN_COMMENTS_NS);
const NAME_CM_LST = qname('p188', 'cmLst', MODERN_COMMENTS_NS);
const NAME_CM = qname('p188', 'cm', MODERN_COMMENTS_NS);
const NAME_REPLY_LST = qname('p188', 'replyLst', MODERN_COMMENTS_NS);
const NAME_REPLY = qname('p188', 'reply', MODERN_COMMENTS_NS);
const NAME_TX_BODY = qname('p188', 'txBody', MODERN_COMMENTS_NS);
const NAME_POS = qname('p188', 'pos', MODERN_COMMENTS_NS);
const NAME_UNKNOWN_ANCHOR = qname('p188', 'unknownAnchor', MODERN_COMMENTS_NS);
const NAME_SLD_MK_LST = qname('pc', 'sldMkLst', PC_NS);
const NAME_DOC_MK = qname('pc', 'docMk', PC_NS);
const NAME_SLD_MK = qname('pc', 'sldMk', PC_NS);
const NAME_BODY_PR = qname('a', 'bodyPr', NS.dml);

const ATTR_ID = qname('', 'id', '');
const ATTR_NAME = qname('', 'name', '');
const ATTR_INITIALS = qname('', 'initials', '');
const ATTR_USER_ID = qname('', 'userId', '');
const ATTR_PROVIDER_ID = qname('', 'providerId', '');
const ATTR_AUTHOR_ID = qname('', 'authorId', '');
const ATTR_STATUS = qname('', 'status', '');
const ATTR_CREATED = qname('', 'created', '');
const ATTR_X = qname('', 'x', '');
const ATTR_Y = qname('', 'y', '');
const ATTR_C_ID = qname('', 'cId', '');
const ATTR_SLD_ID = qname('', 'sldId', '');

// Types --------------------------------------------------------------------

/** `ST_CommentStatus` — the three states a thread can be in. */
export type CommentStatus = 'active' | 'resolved' | 'closed';
export const COMMENT_STATUSES: readonly CommentStatus[] = ['active', 'resolved', 'closed'];

export interface ModernAuthor {
  /** A GUID, braced and upper-case (`ST_AuthorId`). */
  readonly id: string;
  readonly name: string;
  readonly initials: string | null;
  /** Identity within `providerId`'s directory. Required by the schema. */
  readonly userId: string;
  /** Which directory `userId` belongs to. Required by the schema. */
  readonly providerId: string;
}

export interface ModernCommentPosition {
  readonly x: number;
  readonly y: number;
}

/** One reply. Replies carry the same properties a thread does, minus a thread's own. */
export interface ModernReply {
  readonly id: string;
  readonly authorId: string;
  /** `null` when the file states a status this schema version does not define. */
  readonly status: CommentStatus | null;
  readonly created: string | null;
  readonly text: string;
}

export interface ModernComment extends ModernReply {
  readonly replies: ReadonlyArray<ModernReply>;
  readonly position: ModernCommentPosition | null;
}

// Read ---------------------------------------------------------------------

const readStatus = (el: XmlElement): CommentStatus | null => {
  const raw = getAttrValue(el, ATTR_STATUS);
  // Absent means `active`: the schema gives the attribute that default.
  if (raw === null) return 'active';
  return COMMENT_STATUSES.find((status) => status === raw) ?? null;
};

const readText = (el: XmlElement): string => {
  const body = firstChildElement(el, NAME_TX_BODY);
  return body === null ? '' : textBodyText(body);
};

const readProperties = (el: XmlElement): ModernReply | null => {
  const id = getAttrValue(el, ATTR_ID);
  const authorId = getAttrValue(el, ATTR_AUTHOR_ID);
  // Both are required by the schema, and both are how everything else finds
  // this comment again. One without them is not addressable, so it is left
  // where it is rather than reported as something the caller can act on.
  if (id === null || authorId === null) return null;
  return {
    id,
    authorId,
    status: readStatus(el),
    created: getAttrValue(el, ATTR_CREATED),
    text: readText(el),
  };
};

export const readModernAuthorList = (root: XmlElement): ReadonlyArray<ModernAuthor> => {
  if (root.name.namespaceURI !== MODERN_COMMENTS_NS || root.name.localName !== 'authorLst') {
    throw new Error(
      `expected <p188:authorLst> root, got <${root.name.prefix}:${root.name.localName}>`,
    );
  }
  const authors: ModernAuthor[] = [];
  for (const el of allChildElements(root, NAME_AUTHOR)) {
    const id = getAttrValue(el, ATTR_ID);
    if (id === null) continue;
    authors.push({
      id,
      name: getAttrValue(el, ATTR_NAME) ?? '',
      initials: getAttrValue(el, ATTR_INITIALS),
      userId: getAttrValue(el, ATTR_USER_ID) ?? '',
      providerId: getAttrValue(el, ATTR_PROVIDER_ID) ?? '',
    });
  }
  return authors;
};

const readPosition = (el: XmlElement): ModernCommentPosition | null => {
  const pos = firstChildElement(el, NAME_POS);
  if (pos === null) return null;
  const x = Number.parseInt(getAttrValue(pos, ATTR_X) ?? '', 10);
  const y = Number.parseInt(getAttrValue(pos, ATTR_Y) ?? '', 10);
  return Number.isFinite(x) && Number.isFinite(y) ? { x, y } : null;
};

export const readModernCommentList = (root: XmlElement): ReadonlyArray<ModernComment> => {
  if (root.name.namespaceURI !== MODERN_COMMENTS_NS || root.name.localName !== 'cmLst') {
    throw new Error(`expected <p188:cmLst> root, got <${root.name.prefix}:${root.name.localName}>`);
  }
  const comments: ModernComment[] = [];
  for (const el of allChildElements(root, NAME_CM)) {
    const properties = readProperties(el);
    if (properties === null) continue;
    const replyLst = firstChildElement(el, NAME_REPLY_LST);
    const replies: ModernReply[] = [];
    if (replyLst !== null) {
      for (const reply of allChildElements(replyLst, NAME_REPLY)) {
        const read = readProperties(reply);
        if (read !== null) replies.push(read);
      }
    }
    comments.push({ ...properties, replies, position: readPosition(el) });
  }
  return comments;
};

// Find ---------------------------------------------------------------------

/** The `<p188:cm>` with this id, or `null`. */
export const findModernComment = (root: XmlElement, id: string): XmlElement | null =>
  allChildElements(root, NAME_CM).find((el) => getAttrValue(el, ATTR_ID) === id) ?? null;

/** The `<p188:reply>` with this id under `comment`, or `null`. */
export const findModernReply = (comment: XmlElement, id: string): XmlElement | null => {
  const replyLst = firstChildElement(comment, NAME_REPLY_LST);
  if (replyLst === null) return null;
  return (
    allChildElements(replyLst, NAME_REPLY).find((el) => getAttrValue(el, ATTR_ID) === id) ?? null
  );
};

// Build --------------------------------------------------------------------

const buildTextBodyElement = (text: string): XmlElement => {
  // `CT_TextBody` requires `<a:bodyPr>`; the paragraphs are filled in the way
  // every other text body in this library is.
  const body = elem(NAME_TX_BODY, {
    prefixDecls: new Map([['a', NS.dml]]),
    children: [elem(NAME_BODY_PR)],
  });
  setTextBody(body, text);
  return body;
};

// Edit ---------------------------------------------------------------------

/**
 * Sets the status of a thread. `active` is written out rather than left to
 * the schema's default, so a file that has been reopened says so instead of
 * looking like one nobody ever resolved.
 */
export const setModernStatus = (el: XmlElement, status: CommentStatus): void => {
  const existing = el.attrs.findIndex((a) => a.name.localName === 'status' && !a.name.namespaceURI);
  if (existing < 0) el.attrs.push(attr(ATTR_STATUS, status));
  else el.attrs[existing] = { ...el.attrs[existing]!, value: status };
};

/**
 * Replaces the text of a comment or reply, keeping the body's own formatting
 * the way `setShapeText` does. A comment whose `<p188:txBody>` is missing —
 * the schema makes it optional — gets one.
 */
export const setModernText = (el: XmlElement, text: string): void => {
  let body = firstChildElement(el, NAME_TX_BODY);
  if (body === null) {
    body = buildTextBodyElement(text);
    // The sequence puts `txBody` after the anchor, `pos` and `replyLst`, and
    // before `extLst` — which is the only thing that can follow it.
    const extLst = el.children.findIndex(
      (node) =>
        node.kind === 'element' &&
        node.name.namespaceURI === MODERN_COMMENTS_NS &&
        node.name.localName === 'extLst',
    );
    el.children.splice(extLst < 0 ? el.children.length : extLst, 0, body);
    return;
  }
  setTextBody(body, text);
};

// Build --------------------------------------------------------------------

export const buildModernAuthorElement = (author: ModernAuthor): XmlElement => {
  const attrs = [attr(ATTR_ID, author.id), attr(ATTR_NAME, author.name)];
  if (author.initials !== null) attrs.push(attr(ATTR_INITIALS, author.initials));
  attrs.push(attr(ATTR_USER_ID, author.userId), attr(ATTR_PROVIDER_ID, author.providerId));
  return elem(NAME_AUTHOR, { attrs });
};

export const buildModernAuthorListDoc = (authors: ReadonlyArray<ModernAuthor>): XmlDocument => ({
  kind: 'document',
  decl: { version: '1.0', encoding: 'UTF-8', standalone: 'yes' },
  prolog: [],
  root: elem(NAME_AUTHOR_LST, {
    prefixDecls: new Map([['p188', MODERN_COMMENTS_NS]]),
    children: authors.map(buildModernAuthorElement),
  }),
  epilog: [],
});

/**
 * The anchor that says which slide a comment belongs to. The relationship
 * from the slide part already says it, but the schema wants the moniker as
 * well, and it is what PowerPoint matches on.
 */
export const buildSlideAnchor = (creationId: number, slideId: number): XmlElement =>
  elem(NAME_SLD_MK_LST, {
    prefixDecls: new Map([['pc', PC_NS]]),
    children: [
      elem(NAME_DOC_MK),
      elem(NAME_SLD_MK, {
        attrs: [attr(ATTR_C_ID, String(creationId)), attr(ATTR_SLD_ID, String(slideId))],
      }),
    ],
  });

/** For a comment whose slide cannot be identified — the schema's own escape hatch. */
export const buildUnknownAnchor = (): XmlElement => elem(NAME_UNKNOWN_ANCHOR);

export interface ModernCommentSpec {
  readonly id: string;
  readonly authorId: string;
  readonly created: string;
  readonly text: string;
  readonly status?: CommentStatus;
  readonly anchor: XmlElement;
  readonly position?: ModernCommentPosition | null;
}

export const buildModernCommentElement = (spec: ModernCommentSpec): XmlElement => {
  const children: XmlElement[] = [spec.anchor];
  if (spec.position) {
    children.push(
      elem(NAME_POS, {
        attrs: [attr(ATTR_X, String(spec.position.x)), attr(ATTR_Y, String(spec.position.y))],
      }),
    );
  }
  children.push(buildTextBodyElement(spec.text));
  return elem(NAME_CM, {
    prefixDecls: new Map([['p188', MODERN_COMMENTS_NS]]),
    attrs: [
      attr(ATTR_ID, spec.id),
      attr(ATTR_AUTHOR_ID, spec.authorId),
      attr(ATTR_CREATED, spec.created),
      attr(ATTR_STATUS, spec.status ?? 'active'),
    ],
    children,
  });
};

export const buildModernReplyElement = (spec: {
  readonly id: string;
  readonly authorId: string;
  readonly created: string;
  readonly text: string;
}): XmlElement =>
  elem(NAME_REPLY, {
    attrs: [
      attr(ATTR_ID, spec.id),
      attr(ATTR_AUTHOR_ID, spec.authorId),
      attr(ATTR_CREATED, spec.created),
      attr(ATTR_STATUS, 'active'),
    ],
    children: [buildTextBodyElement(spec.text)],
  });

/**
 * Appends a reply to a thread, creating the `<p188:replyLst>` when the thread
 * has none. The list sits between the anchor and the thread's own text body.
 */
export const appendModernReply = (comment: XmlElement, reply: XmlElement): void => {
  const existing = firstChildElement(comment, NAME_REPLY_LST);
  if (existing !== null) {
    existing.children.push(reply);
    return;
  }
  const replyLst = elem(NAME_REPLY_LST, { children: [reply] });
  const body = comment.children.findIndex(
    (node) =>
      node.kind === 'element' &&
      node.name.namespaceURI === MODERN_COMMENTS_NS &&
      (node.name.localName === 'txBody' || node.name.localName === 'extLst'),
  );
  comment.children.splice(body < 0 ? comment.children.length : body, 0, replyLst);
};

export const buildModernCommentListDoc = (comments: ReadonlyArray<XmlElement>): XmlDocument => ({
  kind: 'document',
  decl: { version: '1.0', encoding: 'UTF-8', standalone: 'yes' },
  prolog: [],
  root: elem(NAME_CM_LST, {
    prefixDecls: new Map([
      ['p188', MODERN_COMMENTS_NS],
      ['a', NS.dml],
    ]),
    children: [...comments],
  }),
  epilog: [],
});

/** Removes a thread, or a single reply, by id. Returns whether anything went. */
export const removeModernComment = (root: XmlElement, id: string): boolean => {
  const before = root.children.length;
  root.children = root.children.filter(
    (node) =>
      !(
        node.kind === 'element' &&
        node.name.namespaceURI === MODERN_COMMENTS_NS &&
        node.name.localName === 'cm' &&
        getAttrValue(node, ATTR_ID) === id
      ),
  );
  if (root.children.length !== before) return true;
  for (const comment of allChildElements(root, NAME_CM)) {
    const replyLst = firstChildElement(comment, NAME_REPLY_LST);
    if (replyLst === null) continue;
    const had = replyLst.children.length;
    replyLst.children = replyLst.children.filter(
      (node) =>
        !(
          node.kind === 'element' &&
          node.name.namespaceURI === MODERN_COMMENTS_NS &&
          node.name.localName === 'reply' &&
          getAttrValue(node, ATTR_ID) === id
        ),
    );
    if (replyLst.children.length === had) continue;
    // An empty `<p188:replyLst>` is schema-valid, but PowerPoint does not
    // write one, so neither do we once the last reply goes.
    if (allChildElements(replyLst, NAME_REPLY).length === 0) {
      comment.children = comment.children.filter((node) => node !== replyLst);
    }
    return true;
  }
  return false;
};

/** Every `<p188:cm>` in the list, for callers that mutate in place. */
export const modernCommentElements = (root: XmlElement): ReadonlyArray<XmlElement> =>
  allChildElements(root, NAME_CM);

/** The anchor a thread carries, for reuse by another thread on the same slide. */
export const anchorOf = (comment: XmlElement): XmlElement | null => {
  for (const child of comment.children) {
    if (child.kind !== 'element') continue;
    if (child.name.namespaceURI === PC_NS && child.name.localName === 'sldMkLst') return child;
    if (child.name.namespaceURI === MODERN_COMMENTS_NS && child.name.localName === 'unknownAnchor')
      return child;
    // `ac:deMkLst` / `ac:txMkLst` anchor a comment to a shape or a text range.
    // They name the shape, so they are never reused by a different thread.
  }
  return null;
};
