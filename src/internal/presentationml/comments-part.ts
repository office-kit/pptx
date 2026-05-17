// Comments part — `/ppt/commentAuthors.xml` (one per package) and
// `/ppt/comments/comment{N}.xml` (one per slide that has comments).
//
// Implements the **legacy** comment schema (ECMA-376 Part 1 §19.4),
// which every PowerPoint / Keynote / Google Slides / LibreOffice
// Impress consumer in the wild understands. The modern (`p15:`)
// threaded-comments schema is a strict superset; preserving an
// existing modern-comments part on round-trip is handled separately
// via OPC pass-through.
//
// Schema (legacy):
//
//   <p:cmAuthorLst>
//     <p:cmAuthor id="0" name="Reviewer" initials="R" lastIdx="1" clrIdx="0"/>
//   </p:cmAuthorLst>
//
//   <p:cmLst>
//     <p:cm authorId="0" dt="2026-05-15T12:00:00Z" idx="1">
//       <p:pos x="..." y="..."/>      <!-- optional -->
//       <p:text>Comment text.</p:text>
//     </p:cm>
//   </p:cmLst>
//
// Position coordinates are EMUs (ECMA Part 1 §19.7.2 `ST_Coordinate`).
// `idx` is per-author; `lastIdx` on the author tracks the last used.

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
  text as textNode,
} from '../xml/index.ts';

// QNames -------------------------------------------------------------------

const NAME_CM_AUTHOR_LST = qname('p', 'cmAuthorLst', NS.pml);
const NAME_CM_AUTHOR = qname('p', 'cmAuthor', NS.pml);
const NAME_CM_LST = qname('p', 'cmLst', NS.pml);
const NAME_CM = qname('p', 'cm', NS.pml);
const NAME_POS = qname('p', 'pos', NS.pml);
const NAME_TEXT = qname('p', 'text', NS.pml);

const ATTR_ID = qname('', 'id', '');
const ATTR_NAME = qname('', 'name', '');
const ATTR_INITIALS = qname('', 'initials', '');
const ATTR_LAST_IDX = qname('', 'lastIdx', '');
const ATTR_CLR_IDX = qname('', 'clrIdx', '');
const ATTR_AUTHOR_ID = qname('', 'authorId', '');
const ATTR_DT = qname('', 'dt', '');
const ATTR_IDX = qname('', 'idx', '');
const ATTR_X = qname('', 'x', '');
const ATTR_Y = qname('', 'y', '');

// Types --------------------------------------------------------------------

export interface CommentAuthor {
  /** Unique within the author list (also referenced by `<p:cm authorId="…">`). */
  readonly id: number;
  readonly name: string;
  readonly initials: string;
  /** Highest `idx` value this author has used; bumped on every new comment. */
  readonly lastIdx: number;
  /** Optional color index into the package's theme; preserved verbatim. */
  readonly clrIdx: number | null;
}

export interface CommentPosition {
  readonly x: number;
  readonly y: number;
}

export interface SlideComment {
  /** Author id (matches `CommentAuthor.id`). */
  readonly authorId: number;
  /** Per-author monotonic index. */
  readonly idx: number;
  /** ISO-8601 timestamp string as it appeared on disk; may be `null`. */
  readonly dt: string | null;
  readonly text: string;
  readonly position: CommentPosition | null;
}

export interface CommentAuthorList {
  readonly authors: ReadonlyArray<CommentAuthor>;
}

export interface CommentList {
  readonly comments: ReadonlyArray<SlideComment>;
}

// Read ---------------------------------------------------------------------

export const readCommentAuthorList = (root: XmlElement): CommentAuthorList => {
  if (root.name.namespaceURI !== NS.pml || root.name.localName !== 'cmAuthorLst') {
    throw new Error(
      `expected <p:cmAuthorLst> root, got <${root.name.prefix}:${root.name.localName}>`,
    );
  }
  const authors: CommentAuthor[] = [];
  for (const el of allChildElements(root, NAME_CM_AUTHOR)) {
    const idRaw = getAttrValue(el, ATTR_ID);
    const nameVal = getAttrValue(el, ATTR_NAME);
    const initialsVal = getAttrValue(el, ATTR_INITIALS);
    const lastIdxRaw = getAttrValue(el, ATTR_LAST_IDX);
    const clrIdxRaw = getAttrValue(el, ATTR_CLR_IDX);
    if (idRaw === null) continue;
    const id = Number.parseInt(idRaw, 10);
    if (!Number.isFinite(id)) continue;
    authors.push({
      id,
      name: nameVal ?? '',
      initials: initialsVal ?? '',
      lastIdx: lastIdxRaw !== null ? Number.parseInt(lastIdxRaw, 10) || 0 : 0,
      clrIdx: clrIdxRaw !== null ? Number.parseInt(clrIdxRaw, 10) || null : null,
    });
  }
  return { authors };
};

export const readCommentList = (root: XmlElement): CommentList => {
  if (root.name.namespaceURI !== NS.pml || root.name.localName !== 'cmLst') {
    throw new Error(`expected <p:cmLst> root, got <${root.name.prefix}:${root.name.localName}>`);
  }
  const comments: SlideComment[] = [];
  for (const el of allChildElements(root, NAME_CM)) {
    const authorIdRaw = getAttrValue(el, ATTR_AUTHOR_ID);
    const idxRaw = getAttrValue(el, ATTR_IDX);
    const dtVal = getAttrValue(el, ATTR_DT);
    if (authorIdRaw === null || idxRaw === null) continue;
    const authorId = Number.parseInt(authorIdRaw, 10);
    const idx = Number.parseInt(idxRaw, 10);
    if (!Number.isFinite(authorId) || !Number.isFinite(idx)) continue;

    let position: CommentPosition | null = null;
    const posEl = firstChildElement(el, NAME_POS);
    if (posEl !== null) {
      const xRaw = getAttrValue(posEl, ATTR_X);
      const yRaw = getAttrValue(posEl, ATTR_Y);
      if (xRaw !== null && yRaw !== null) {
        const x = Number.parseInt(xRaw, 10);
        const y = Number.parseInt(yRaw, 10);
        if (Number.isFinite(x) && Number.isFinite(y)) {
          position = { x, y };
        }
      }
    }

    let text = '';
    const textEl = firstChildElement(el, NAME_TEXT);
    if (textEl !== null) {
      for (const child of textEl.children) {
        if (child.kind === 'text' || child.kind === 'cdata') text += child.data;
      }
    }

    comments.push({
      authorId,
      idx,
      dt: dtVal,
      text,
      position,
    });
  }
  return { comments };
};

// Build --------------------------------------------------------------------

const commentAuthorElement = (a: CommentAuthor): XmlElement => {
  // `clrIdx` is required by the strict ECMA-376 schema even though
  // PowerPoint tolerates its absence in practice. Default to 0 so we
  // always pass `xmllint --schema` validation.
  const attrs = [
    attr(ATTR_ID, String(a.id)),
    attr(ATTR_NAME, a.name),
    attr(ATTR_INITIALS, a.initials),
    attr(ATTR_LAST_IDX, String(a.lastIdx)),
    attr(ATTR_CLR_IDX, String(a.clrIdx ?? 0)),
  ];
  return elem(NAME_CM_AUTHOR, { attrs });
};

export const buildCommentAuthorListDoc = (authors: ReadonlyArray<CommentAuthor>): XmlDocument => {
  const root = elem(NAME_CM_AUTHOR_LST, {
    prefixDecls: new Map([['p', NS.pml]]),
    children: authors.map(commentAuthorElement),
  });
  return {
    kind: 'document',
    decl: { version: '1.0', encoding: 'UTF-8', standalone: 'yes' },
    prolog: [],
    root,
    epilog: [],
  };
};

const commentElement = (c: SlideComment): XmlElement => {
  const attrs = [attr(ATTR_AUTHOR_ID, String(c.authorId)), attr(ATTR_IDX, String(c.idx))];
  if (c.dt !== null) attrs.push(attr(ATTR_DT, c.dt));
  const children: XmlElement[] = [];
  if (c.position !== null) {
    children.push(
      elem(NAME_POS, {
        attrs: [attr(ATTR_X, String(c.position.x)), attr(ATTR_Y, String(c.position.y))],
      }),
    );
  }
  children.push(elem(NAME_TEXT, { children: [textNode(c.text)] }));
  return elem(NAME_CM, { attrs, children });
};

export const buildCommentListDoc = (comments: ReadonlyArray<SlideComment>): XmlDocument => {
  const root = elem(NAME_CM_LST, {
    prefixDecls: new Map([['p', NS.pml]]),
    children: comments.map(commentElement),
  });
  return {
    kind: 'document',
    decl: { version: '1.0', encoding: 'UTF-8', standalone: 'yes' },
    prolog: [],
    root,
    epilog: [],
  };
};
