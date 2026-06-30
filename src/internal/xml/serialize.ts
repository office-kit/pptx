// XML serializer for the AST in `./ast.ts`.
//
// Goals:
//
//   - Round-trip fidelity. Element/attribute order and prefix choices on the
//     way out match the way in (provided the AST hasn't been mutated to break
//     that). The serializer never reorders attributes, never reallocates
//     namespace prefixes, never re-emits xmlns declarations on a different
//     element than where the parser placed them.
//   - Predictable output for hand-built ASTs. Authoring-time code can assume
//     the same input element produces the same byte sequence.
//   - Escape rules tight enough for PowerPoint to accept the output: the
//     predefined entities, numeric references for the whitespace controls that
//     attribute-value normalization would otherwise eat (tab / LF / CR), and a
//     hard reject for the C0 control characters XML 1.0 forbids outright.
//
// Out of scope:
//
//   - Pretty-printing. OOXML files are usually compact-on-write because most
//     producers don't pretty-print. A future option could add indent if
//     someone really wants it, but the round-trip story is cleaner without.

import type { XmlAttr, XmlDocument, XmlElement, XmlNode } from './ast.ts';

// XML 1.0 forbids the C0 control characters except tab (0x09), LF (0x0A), and
// CR (0x0D). Unlike most escapable characters these cannot be rescued with a
// numeric reference either — `&#0;` … `&#8;` are themselves illegal because the
// referenced code point is not a legal `Char` (XML 1.0 §2.2). A value carrying
// one can only have entered through authoring input (a parsed document could not
// have held it), so we reject loudly: emitting it raw produces a non-well-formed
// part that corrupts the entire .pptx, and silently dropping it would be a
// surprise data loss. Strip the character upstream before authoring.
const rejectForbiddenControlChar = (c: number): void => {
  if (c < 0x20 && c !== 0x09 && c !== 0x0a && c !== 0x0d) {
    const hex = c.toString(16).toUpperCase().padStart(4, '0');
    throw new Error(
      `XML-illegal control character U+${hex} in text; strip control characters before authoring`,
    );
  }
};

const escapeAttr = (s: string): string => {
  // Attribute values: escape <, &, and the quote char we use. We always use
  // double quotes, so single quotes pass through. CR and LF must be encoded
  // numerically because XML parsers normalize whitespace in attribute values.
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    rejectForbiddenControlChar(c);
    if (c === 38) out += '&amp;';
    else if (c === 60) out += '&lt;';
    else if (c === 62) out += '&gt;';
    else if (c === 34) out += '&quot;';
    else if (c === 9) out += '&#9;';
    else if (c === 10) out += '&#10;';
    else if (c === 13) out += '&#13;';
    else out += s[i];
  }
  return out;
};

const escapeText = (s: string): string => {
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    rejectForbiddenControlChar(c);
    if (c === 38) out += '&amp;';
    else if (c === 60) out += '&lt;';
    else if (c === 62) out += '&gt;';
    else if (c === 13) out += '&#13;';
    else out += s[i];
  }
  return out;
};

const rawName = (prefix: string, localName: string): string =>
  prefix === '' ? localName : `${prefix}:${localName}`;

const writeAttr = (a: XmlAttr): string => {
  const name = rawName(a.name.prefix, a.name.localName);
  return ` ${name}="${escapeAttr(a.value)}"`;
};

const writeElement = (e: XmlElement, parts: string[]): void => {
  const name = rawName(e.name.prefix, e.name.localName);
  parts.push('<');
  parts.push(name);
  for (const [prefix, uri] of e.prefixDecls) {
    const decl = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
    parts.push(` ${decl}="${escapeAttr(uri)}"`);
  }
  for (const a of e.attrs) parts.push(writeAttr(a));
  if (e.children.length === 0) {
    parts.push('/>');
    return;
  }
  parts.push('>');
  for (const child of e.children) writeNode(child, parts);
  parts.push('</');
  parts.push(name);
  parts.push('>');
};

const writeNode = (n: XmlNode, parts: string[]): void => {
  switch (n.kind) {
    case 'element':
      writeElement(n, parts);
      return;
    case 'text':
      parts.push(escapeText(n.data));
      return;
    case 'cdata':
      // CDATA may not contain `]]>`. The canonical workaround is to break the
      // illegal sequence by ending one CDATA section and starting another:
      // replace `]]>` with `]]]]><![CDATA[>` (which re-parses as `]]` followed
      // by `>`). The result is two adjacent CDATA sections whose concatenated
      // content equals the original. PowerPoint never emits CDATA in PPTX, so
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
    parts.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  }
  for (const p of doc.prolog) writeNode(p, parts);
  writeElement(doc.root, parts);
  for (const e of doc.epilog) writeNode(e, parts);
  return parts.join('');
};

// Serialize a single element with no declaration / document wrapper. Useful in
// tests and when emitting fragments that get inlined into a larger document.
export const serializeFragment = (element: XmlElement): string => {
  const parts: string[] = [];
  writeElement(element, parts);
  return parts.join('');
};

export const _internalsForTest = { escapeAttr, escapeText };
