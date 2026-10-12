// XML serializer for the AST in `./ast.ts`.
//
// Goals:
//
//   - Round-trip fidelity. Element/attribute order and prefix choices on the
//     way out match the way in (provided the AST hasn't been mutated to break
//     that). The serializer never reorders attributes, never reallocates
//     namespace prefixes, never re-emits xmlns declarations on a different
//     element than where the parser placed them. It only adds a declaration
//     where an element's or attribute's prefix would otherwise be unbound.
//   - Predictable output for hand-built ASTs. Authoring-time code can assume
//     the same input element produces the same byte sequence.
//   - Escape rules tight enough for the reference desktop app to accept the output: the
//     predefined entities, numeric references for the whitespace controls that
//     attribute-value normalization would otherwise eat (tab / LF / CR), and a
//     hard reject for characters outside the XML 1.0 Char production.
//
// Out of scope:
//
//   - Pretty-printing. OOXML files are usually compact-on-write because most
//     producers don't pretty-print. A future option could add indent if
//     someone really wants it, but the round-trip story is cleaner without.

import type { XmlAttr, XmlDocument, XmlElement, XmlNode } from './ast.ts';
import { NS } from './namespaces.ts';

// XML 1.0 §2.2 excludes C0 controls other than tab / LF / CR, U+FFFE,
// U+FFFF, and unpaired surrogates. Numeric references cannot rescue them.
// Iterate by code point so valid surrogate pairs remain legal supplementary
// characters, while lone surrogates are rejected before UTF-8 replaces them.
const rejectForbiddenChar = (c: number): void => {
  const control = c < 0x20 && c !== 0x09 && c !== 0x0a && c !== 0x0d;
  if (control || (c >= 0xd800 && c <= 0xdfff) || c === 0xfffe || c === 0xffff) {
    const hex = c.toString(16).toUpperCase().padStart(4, '0');
    const kind = control ? 'control character' : 'character';
    throw new Error(
      `XML-illegal ${kind} U+${hex} in text; strip ${control ? 'control' : 'illegal'} characters before authoring`,
    );
  }
};

const escapeAttr = (s: string): string => {
  // Attribute values: escape <, &, and the quote char we use. We always use
  // double quotes, so single quotes pass through. CR and LF must be encoded
  // numerically because XML parsers normalize whitespace in attribute values.
  let out = '';
  for (const char of s) {
    const c = char.codePointAt(0)!;
    rejectForbiddenChar(c);
    if (c === 38) out += '&amp;';
    else if (c === 60) out += '&lt;';
    else if (c === 62) out += '&gt;';
    else if (c === 34) out += '&quot;';
    else if (c === 9) out += '&#9;';
    else if (c === 10) out += '&#10;';
    else if (c === 13) out += '&#13;';
    else out += char;
  }
  return out;
};

const escapeText = (s: string): string => {
  let out = '';
  for (const char of s) {
    const c = char.codePointAt(0)!;
    rejectForbiddenChar(c);
    if (c === 38) out += '&amp;';
    else if (c === 60) out += '&lt;';
    else if (c === 62) out += '&gt;';
    else if (c === 13) out += '&#13;';
    else out += char;
  }
  return out;
};

const rawName = (prefix: string, localName: string): string =>
  prefix === '' ? localName : `${prefix}:${localName}`;

const writeAttr = (a: XmlAttr): string => {
  const name = rawName(a.name.prefix, a.name.localName);
  return ` ${name}="${escapeAttr(a.value)}"`;
};

// Prefix → URI bindings in scope at the element being written. The root of a
// document starts with only the reserved `xml` prefix bound.
type Scope = ReadonlyMap<string, string>;

const DOCUMENT_SCOPE: Scope = new Map([['xml', NS.xml]]);

// Parsed elements carry the declarations their input had, but elements built
// with `elem(...)` carry none and rely on an ancestor declaring their prefixes.
// Producers may declare a prefix on the element that uses it rather than on
// the part root, so that ancestor is not guaranteed; declare whatever the
// element's own name and attributes need that is not already in scope.
const missingDecls = (e: XmlElement, scope: Scope): Map<string, string> | null => {
  let added: Map<string, string> | null = null;
  for (let i = -1; i < e.attrs.length; i++) {
    const name = i < 0 ? e.name : e.attrs[i]!.name;
    // Unprefixed attributes are in no namespace whatever the default is.
    if (i >= 0 && name.prefix === '') continue;
    const own = e.prefixDecls.get(name.prefix) ?? added?.get(name.prefix);
    if (own === undefined) {
      if ((scope.get(name.prefix) ?? '') === name.namespaceURI) continue;
      added ??= new Map();
      added.set(name.prefix, name.namespaceURI);
    } else if (own !== name.namespaceURI) {
      throw new Error(
        `prefix "${name.prefix}" on <${rawName(e.name.prefix, e.name.localName)}> is bound to both ${own} and ${name.namespaceURI}`,
      );
    }
  }
  return added;
};

const writeDecls = (decls: ReadonlyMap<string, string>, parts: string[]): void => {
  for (const [prefix, uri] of decls) {
    const decl = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
    parts.push(` ${decl}="${escapeAttr(uri)}"`);
  }
};

const withDecls = (scope: Scope, decls: ReadonlyMap<string, string>): Scope => {
  if (decls.size === 0) return scope;
  const next = new Map(scope);
  for (const [prefix, uri] of decls) next.set(prefix, uri);
  return next;
};

// `scope` is null for fragments: their context is unknown, so they are written
// with exactly the declarations they carry.
const writeElement = (e: XmlElement, parts: string[], scope: Scope | null): void => {
  const name = rawName(e.name.prefix, e.name.localName);
  parts.push('<');
  parts.push(name);
  writeDecls(e.prefixDecls, parts);
  let childScope = scope;
  if (scope) {
    const added = missingDecls(e, scope);
    childScope = withDecls(scope, e.prefixDecls);
    if (added) {
      writeDecls(added, parts);
      childScope = withDecls(childScope, added);
    }
  }
  for (const a of e.attrs) parts.push(writeAttr(a));
  if (e.children.length === 0) {
    parts.push('/>');
    return;
  }
  parts.push('>');
  for (const child of e.children) writeNode(child, parts, childScope);
  parts.push('</');
  parts.push(name);
  parts.push('>');
};

const writeNode = (n: XmlNode, parts: string[], scope: Scope | null): void => {
  switch (n.kind) {
    case 'element':
      writeElement(n, parts, scope);
      return;
    case 'text':
      parts.push(escapeText(n.data));
      return;
    case 'cdata':
      // CDATA may not contain `]]>`. The canonical workaround is to break the
      // illegal sequence by ending one CDATA section and starting another:
      // replace `]]>` with `]]]]><![CDATA[>` (which re-parses as `]]` followed
      // by `>`). The result is two adjacent CDATA sections whose concatenated
      // content equals the original. The reference desktop app never emits CDATA in PPTX, so
      // this path only ever runs for hand-built ASTs.
      parts.push('<![CDATA[');
      parts.push(n.data.split(']]>').join(']]]]><![CDATA[>'));
      parts.push(']]>');
      return;
    case 'comment':
      parts.push('<!--');
      parts.push(n.data);
      parts.push('-->');
      return;
    case 'pi':
      parts.push('<?');
      parts.push(n.target);
      if (n.data.length > 0) {
        parts.push(' ');
        parts.push(n.data);
      }
      parts.push('?>');
      return;
  }
};

export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const writeDeclaration = (doc: XmlDocument, parts: string[]): void => {
  if (!doc.decl) return;
  parts.push('<?xml version="');
  parts.push(escapeAttr(doc.decl.version));
  parts.push('"');
  if (doc.decl.encoding !== undefined) {
    parts.push(' encoding="');
    parts.push(escapeAttr(doc.decl.encoding));
    parts.push('"');
  }
  if (doc.decl.standalone !== undefined) {
    parts.push(' standalone="');
    parts.push(doc.decl.standalone);
    parts.push('"');
  }
  parts.push('?>');
};

export const serializeXml = (
  doc: XmlDocument,
  options: { standaloneDeclaration?: boolean } = {},
): string => {
  const parts: string[] = [];
  if (doc.decl) {
    writeDeclaration(doc, parts);
  } else if (options.standaloneDeclaration ?? true) {
    // OOXML parts MUST start with an XML declaration. Add a sensible default
    // when the AST doesn't carry one (typical for hand-built trees).
    parts.push(XML_DECLARATION);
  }
  for (const p of doc.prolog) writeNode(p, parts, null);
  writeElement(doc.root, parts, DOCUMENT_SCOPE);
  for (const e of doc.epilog) writeNode(e, parts, null);
  return parts.join('');
};

// Serialize a single element with no declaration / document wrapper. Useful in
// tests and when emitting fragments that get inlined into a larger document.
export const serializeFragment = (element: XmlElement): string => {
  const parts: string[] = [];
  writeElement(element, parts, null);
  return parts.join('');
};

export const _internalsForTest = { escapeAttr, escapeText };
