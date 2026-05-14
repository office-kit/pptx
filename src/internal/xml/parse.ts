// XML parser tuned for OOXML.
//
// Hand-written tokenizer + tree builder. Handles the subset of XML 1.0 that
// PowerPoint and the OOXML toolchain actually emit:
//
//   - The XML declaration `<?xml ... ?>`.
//   - Element start / end / empty tags with attributes.
//   - Namespace prefix declarations (`xmlns`, `xmlns:foo`).
//   - Character data, with the five predefined entities and numeric character
//     references.
//   - CDATA sections, comments, processing instructions.
//
// Deliberately NOT supported:
//
//   - DOCTYPE declarations and DTDs — OOXML doesn't use them. We skip past a
//     leading DOCTYPE if present, but its internal subset is rejected.
//   - General entity references beyond the five predefined names.
//   - XML 1.1 line-end normalization differences.
//
// The parser throws `XmlParseError` with a position on the first malformed
// construct. It is intentionally strict: a silent recovery here would mask
// data corruption further down the pipeline.

import {
  type QName,
  type XmlAttr,
  type XmlCData,
  type XmlComment,
  type XmlDeclaration,
  type XmlDocument,
  type XmlElement,
  type XmlPI,
  attr,
  cdata,
  comment,
  elem,
  pi,
  qname,
  text,
} from './ast.ts';
import { NS } from './namespaces.ts';

export class XmlParseError extends Error {
  constructor(
    message: string,
    public readonly offset: number,
    public readonly line: number,
    public readonly column: number,
  ) {
    super(`${message} at line ${line}:${column} (offset ${offset})`);
    this.name = 'XmlParseError';
  }
}

interface Cursor {
  readonly src: string;
  pos: number;
}

const isNameStartChar = (c: string): boolean => {
  // XML 1.0 NameStartChar simplified to what OOXML actually emits: ASCII
  // letters, underscore, colon (used in prefixes). The full Unicode predicate
  // is overkill — PowerPoint never emits non-ASCII element names.
  const code = c.charCodeAt(0);
  return (
    (code >= 65 && code <= 90) || // A-Z
    (code >= 97 && code <= 122) || // a-z
    code === 95 || // _
    code === 58 // :
  );
};

const isNameChar = (c: string): boolean => {
  if (isNameStartChar(c)) return true;
  const code = c.charCodeAt(0);
  return (
    (code >= 48 && code <= 57) || // 0-9
    code === 45 || // -
    code === 46 // .
  );
};

const isWhitespace = (c: string): boolean => c === ' ' || c === '\t' || c === '\n' || c === '\r';

const lineColOf = (src: string, pos: number): { line: number; column: number } => {
  let line = 1;
  let column = 1;
  for (let i = 0; i < pos && i < src.length; i++) {
    if (src.charCodeAt(i) === 10) {
      line++;
      column = 1;
    } else {
      column++;
    }
  }
  return { line, column };
};

const fail = (cur: Cursor, message: string): never => {
  const { line, column } = lineColOf(cur.src, cur.pos);
  throw new XmlParseError(message, cur.pos, line, column);
};

const expect = (cur: Cursor, literal: string): void => {
  if (cur.src.slice(cur.pos, cur.pos + literal.length) !== literal) {
    fail(cur, `expected "${literal}"`);
  }
  cur.pos += literal.length;
};

const accept = (cur: Cursor, literal: string): boolean => {
  if (cur.src.slice(cur.pos, cur.pos + literal.length) !== literal) return false;
  cur.pos += literal.length;
  return true;
};

const skipWhitespace = (cur: Cursor): void => {
  while (cur.pos < cur.src.length && isWhitespace(cur.src[cur.pos] ?? '')) cur.pos++;
};

const readName = (cur: Cursor): string => {
  const start = cur.pos;
  if (cur.pos >= cur.src.length || !isNameStartChar(cur.src[cur.pos] ?? '')) {
    fail(cur, 'expected XML name');
  }
  cur.pos++;
  while (cur.pos < cur.src.length && isNameChar(cur.src[cur.pos] ?? '')) cur.pos++;
  return cur.src.slice(start, cur.pos);
};

const splitQName = (raw: string): { prefix: string; localName: string } => {
  const idx = raw.indexOf(':');
  if (idx < 0) return { prefix: '', localName: raw };
  return { prefix: raw.slice(0, idx), localName: raw.slice(idx + 1) };
};

// Character / entity references inside attribute values and text content.
// XML 1.0 predefined entities + decimal/hex numeric character references.
const decodeRefs = (cur: Cursor, raw: string): string => {
  let out = '';
  let i = 0;
  while (i < raw.length) {
    const ch = raw.charCodeAt(i);
    if (ch !== 38 /* & */) {
      out += raw[i];
      i++;
      continue;
    }
    const end = raw.indexOf(';', i + 1);
    if (end < 0) {
      cur.pos -= raw.length - i;
      fail(cur, 'unterminated entity reference');
    }
    const ref = raw.slice(i + 1, end);
    if (ref === 'amp') out += '&';
    else if (ref === 'lt') out += '<';
    else if (ref === 'gt') out += '>';
    else if (ref === 'quot') out += '"';
    else if (ref === 'apos') out += "'";
    else if (ref.startsWith('#x') || ref.startsWith('#X')) {
      const code = Number.parseInt(ref.slice(2), 16);
      if (!Number.isFinite(code)) fail(cur, `invalid hex numeric reference: &${ref};`);
      out += String.fromCodePoint(code);
    } else if (ref.startsWith('#')) {
      const code = Number.parseInt(ref.slice(1), 10);
      if (!Number.isFinite(code)) fail(cur, `invalid decimal numeric reference: &${ref};`);
      out += String.fromCodePoint(code);
    } else {
      fail(cur, `unknown entity reference: &${ref};`);
    }
    i = end + 1;
  }
  return out;
};

const readAttrValue = (cur: Cursor): string => {
  const quote = cur.src[cur.pos];
  if (quote !== '"' && quote !== "'") fail(cur, 'expected attribute value');
  cur.pos++;
  const start = cur.pos;
  while (cur.pos < cur.src.length && cur.src[cur.pos] !== quote) {
    if (cur.src.charCodeAt(cur.pos) === 60 /* < */) {
      fail(cur, '"<" is not allowed inside an attribute value');
    }
    cur.pos++;
  }
  if (cur.pos >= cur.src.length) fail(cur, 'unterminated attribute value');
  const raw = cur.src.slice(start, cur.pos);
  cur.pos++; // consume closing quote
  return decodeRefs(cur, raw);
};

interface NamespaceScope {
  readonly parent: NamespaceScope | null;
  /** prefix → URI; '' key holds the default namespace. */
  readonly bindings: Map<string, string>;
}

const xmlScope: NamespaceScope = {
  parent: null,
  bindings: new Map([['xml', NS.xml]]),
};

const resolvePrefix = (scope: NamespaceScope, prefix: string): string | null => {
  let cursor: NamespaceScope | null = scope;
  while (cursor) {
    const found = cursor.bindings.get(prefix);
    if (found !== undefined) return found;
    cursor = cursor.parent;
  }
  return null;
};

const resolveQName = (
  cur: Cursor,
  scope: NamespaceScope,
  raw: string,
  isAttribute: boolean,
): QName => {
  const { prefix, localName } = splitQName(raw);
  let namespaceURI = '';
  if (prefix !== '') {
    const found = resolvePrefix(scope, prefix);
    if (found === null) {
      return fail(cur, `unbound namespace prefix "${prefix}"`);
    }
    namespaceURI = found;
  } else if (!isAttribute) {
    // Unqualified element names use the default namespace, if any.
    namespaceURI = resolvePrefix(scope, '') ?? '';
  }
  // Per Namespaces in XML 1.0: unqualified attributes have NO namespace, even
  // if a default namespace is declared. We honor that.
  return qname(prefix, localName, namespaceURI);
};

const parseDeclaration = (cur: Cursor): XmlDeclaration | null => {
  if (!accept(cur, '<?xml')) return null;
  // Anything until `?>`
  const end = cur.src.indexOf('?>', cur.pos);
  if (end < 0) fail(cur, 'unterminated XML declaration');
  const body = cur.src.slice(cur.pos, end);
  cur.pos = end + 2;
  const get = (key: string): string | undefined => {
    const m = body.match(new RegExp(`\\b${key}\\s*=\\s*(['"])(.*?)\\1`));
    return m?.[2];
  };
  const version = get('version');
  if (version === undefined) {
    return fail(cur, 'XML declaration missing "version"');
  }
  const encoding = get('encoding');
  const standaloneRaw = get('standalone');
  let standalone: 'yes' | 'no' | undefined;
  if (standaloneRaw === 'yes' || standaloneRaw === 'no') {
    standalone = standaloneRaw;
  } else if (standaloneRaw !== undefined) {
    return fail(cur, `invalid standalone value "${standaloneRaw}"`);
  }
  return {
    version,
    ...(encoding !== undefined ? { encoding } : {}),
    ...(standalone !== undefined ? { standalone } : {}),
  };
};

const parseComment = (cur: Cursor): XmlComment => {
  expect(cur, '<!--');
  const end = cur.src.indexOf('-->', cur.pos);
  if (end < 0) fail(cur, 'unterminated comment');
  const data = cur.src.slice(cur.pos, end);
  if (data.includes('--')) fail(cur, '"--" is not allowed inside an XML comment');
  cur.pos = end + 3;
  return comment(data);
};

const parsePI = (cur: Cursor): XmlPI => {
  expect(cur, '<?');
  const targetStart = cur.pos;
  while (
    cur.pos < cur.src.length &&
    !isWhitespace(cur.src[cur.pos] ?? '') &&
    cur.src.slice(cur.pos, cur.pos + 2) !== '?>'
  ) {
    cur.pos++;
  }
  const target = cur.src.slice(targetStart, cur.pos);
  if (target.toLowerCase() === 'xml') {
    fail(cur, 'PI target may not be "xml"');
  }
  skipWhitespace(cur);
  const end = cur.src.indexOf('?>', cur.pos);
  if (end < 0) fail(cur, 'unterminated processing instruction');
  const data = cur.src.slice(cur.pos, end);
  cur.pos = end + 2;
  return pi(target, data);
};

const parseCData = (cur: Cursor): XmlCData => {
  expect(cur, '<![CDATA[');
  const end = cur.src.indexOf(']]>', cur.pos);
  if (end < 0) fail(cur, 'unterminated CDATA section');
  const data = cur.src.slice(cur.pos, end);
  cur.pos = end + 3;
  return cdata(data);
};

const skipDoctype = (cur: Cursor): void => {
  // OOXML does not use DOCTYPE; if we encounter one we skip it without parsing
  // the internal subset (no `[` allowed).
  expect(cur, '<!DOCTYPE');
  while (cur.pos < cur.src.length) {
    const c = cur.src[cur.pos];
    if (c === '[') fail(cur, 'DOCTYPE internal subset is not supported');
    if (c === '>') {
      cur.pos++;
      return;
    }
    cur.pos++;
  }
  fail(cur, 'unterminated DOCTYPE');
};

const parseStartOrEmptyTag = (
  cur: Cursor,
  parentScope: NamespaceScope,
): { element: XmlElement; scope: NamespaceScope; empty: boolean } => {
  expect(cur, '<');
  const rawName = readName(cur);
  const ownAttrs: { raw: string; rawValue: string }[] = [];
  const prefixDecls = new Map<string, string>();
  let empty = false;

  while (true) {
    const wsStart = cur.pos;
    skipWhitespace(cur);
    if (cur.pos === wsStart && cur.src[cur.pos] !== '/' && cur.src[cur.pos] !== '>') {
      fail(cur, 'expected whitespace before next attribute');
    }
    if (accept(cur, '/>')) {
      empty = true;
      break;
    }
    if (accept(cur, '>')) break;
    const attrRawName = readName(cur);
    skipWhitespace(cur);
    expect(cur, '=');
    skipWhitespace(cur);
    const value = readAttrValue(cur);
    if (attrRawName === 'xmlns') {
      prefixDecls.set('', value);
    } else if (attrRawName.startsWith('xmlns:')) {
      const p = attrRawName.slice(6);
      if (p === '') fail(cur, 'empty namespace prefix in xmlns:');
      if (p === 'xmlns') fail(cur, 'prefix "xmlns" cannot be redeclared');
      prefixDecls.set(p, value);
    } else {
      ownAttrs.push({ raw: attrRawName, rawValue: value });
    }
  }

  const scope: NamespaceScope = {
    parent: parentScope,
    bindings: new Map<string, string>(prefixDecls),
  };

  const elemName = resolveQName(cur, scope, rawName, false);
  const attrs: XmlAttr[] = ownAttrs.map(({ raw, rawValue }) =>
    attr(resolveQName(cur, scope, raw, true), rawValue),
  );

  return {
    element: elem(elemName, { attrs, prefixDecls, children: [] }),
    scope,
    empty,
  };
};

const parseEndTag = (cur: Cursor): string => {
  expect(cur, '</');
  const name = readName(cur);
  skipWhitespace(cur);
  expect(cur, '>');
  return name;
};

const parseTextContent = (cur: Cursor): string => {
  const start = cur.pos;
  while (cur.pos < cur.src.length && cur.src.charCodeAt(cur.pos) !== 60 /* < */) {
    cur.pos++;
  }
  return decodeRefs(cur, cur.src.slice(start, cur.pos));
};

const parseElement = (cur: Cursor, parentScope: NamespaceScope): XmlElement => {
  const { element, scope, empty } = parseStartOrEmptyTag(cur, parentScope);
  if (empty) return element;
  while (cur.pos < cur.src.length) {
    if (cur.src.charCodeAt(cur.pos) === 60 /* < */) {
      if (cur.src.startsWith('</', cur.pos)) {
        const closingRaw = parseEndTag(cur);
        const expectedRaw =
          element.name.prefix === ''
            ? element.name.localName
            : `${element.name.prefix}:${element.name.localName}`;
        if (closingRaw !== expectedRaw) {
          fail(cur, `mismatched closing tag: expected </${expectedRaw}>, got </${closingRaw}>`);
        }
        return element;
      }
      if (cur.src.startsWith('<![CDATA[', cur.pos)) {
        element.children.push(parseCData(cur));
      } else if (cur.src.startsWith('<!--', cur.pos)) {
        element.children.push(parseComment(cur));
      } else if (cur.src.startsWith('<?', cur.pos)) {
        element.children.push(parsePI(cur));
      } else {
        element.children.push(parseElement(cur, scope));
      }
    } else {
      const t = parseTextContent(cur);
      if (t.length > 0) element.children.push(text(t));
    }
  }
  return fail(cur, `unterminated element <${element.name.localName}>`);
};

export const parseXml = (src: string): XmlDocument => {
  // Strip a leading BOM if present — common in OOXML files written by .NET.
  const stripped = src.charCodeAt(0) === 0xfeff ? src.slice(1) : src;
  const cur: Cursor = { src: stripped, pos: 0 };
  const decl = parseDeclaration(cur);
  skipWhitespace(cur);

  const prolog: Array<XmlComment | XmlPI> = [];
  // Prolog: comments, PIs, optional DOCTYPE. Stop at the root element.
  while (cur.pos < cur.src.length) {
    skipWhitespace(cur);
    if (cur.src.startsWith('<!--', cur.pos)) prolog.push(parseComment(cur));
    else if (cur.src.startsWith('<!DOCTYPE', cur.pos)) skipDoctype(cur);
    else if (cur.src.startsWith('<?', cur.pos)) prolog.push(parsePI(cur));
    else break;
  }
  skipWhitespace(cur);

  if (cur.pos >= cur.src.length || cur.src[cur.pos] !== '<') {
    fail(cur, 'expected root element');
  }
  const root = parseElement(cur, xmlScope);

  const epilog: Array<XmlComment | XmlPI> = [];
  while (cur.pos < cur.src.length) {
    skipWhitespace(cur);
    if (cur.pos >= cur.src.length) break;
    if (cur.src.startsWith('<!--', cur.pos)) epilog.push(parseComment(cur));
    else if (cur.src.startsWith('<?', cur.pos)) epilog.push(parsePI(cur));
    else fail(cur, 'unexpected content after root element');
  }

  return { kind: 'document', decl, prolog, root, epilog };
};

// Convenience for parsing an XML fragment (no declaration / prolog), returning
// just the root element. Used heavily by part-specific tests.
export const parseFragment = (src: string): XmlElement => parseXml(src).root;

// Lightweight helper for hand-written test cases: builds a scope chain from
// inherited prefix declarations and resolves a raw "prefix:local" string.
export const resolveQNameWithBindings = (
  raw: string,
  bindings: ReadonlyArray<readonly [prefix: string, uri: string]>,
  isAttribute: boolean,
): QName => {
  const scope: NamespaceScope = {
    parent: null,
    bindings: new Map(bindings),
  };
  return resolveQName({ src: raw, pos: 0 }, scope, raw, isAttribute);
};

export const _internalsForTest = { isNameStartChar, isNameChar, splitQName };
