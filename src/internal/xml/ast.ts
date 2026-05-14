// XML AST types. Designed for round-trip fidelity on OOXML:
//
// - QName carries BOTH the namespace URI (for semantic checks) AND the prefix
//   the input used (so we can serialize back with the same prefix). DOM-style
//   APIs that re-derive prefixes from URIs break PowerPoint's compatibility
//   tooling.
// - Attributes are an ordered array, not a map. OOXML attribute order matters
//   for tools that string-compare serialized XML (Microsoft Open XML Diff).
// - Namespace declarations are recorded on the element where they appeared,
//   not flattened. `mc:AlternateContent` and similar nested namespace scopes
//   would otherwise lose their structure.
// - Text vs CData vs Comment vs PI are distinct node kinds so the serializer
//   doesn't have to guess what to emit.

export interface QName {
  /** Empty string when no prefix is used (default namespace or unqualified). */
  readonly prefix: string;
  readonly localName: string;
  /** Empty string when the name has no namespace binding. */
  readonly namespaceURI: string;
}

export interface XmlAttr {
  readonly name: QName;
  readonly value: string;
}

export interface XmlElement {
  readonly kind: 'element';
  name: QName;
  attrs: XmlAttr[];
  /**
   * Prefix → URI declarations that appeared as `xmlns` / `xmlns:foo`
   * attributes on this element. The serializer re-emits them as attributes
   * (NOT inside `attrs`, since that array is for non-namespace attributes
   * only). An empty-string prefix entry is a default-namespace declaration.
   */
  prefixDecls: Map<string, string>;
  children: XmlNode[];
}

export interface XmlText {
  readonly kind: 'text';
  data: string;
}

export interface XmlCData {
  readonly kind: 'cdata';
  data: string;
}

export interface XmlComment {
  readonly kind: 'comment';
  data: string;
}

export interface XmlPI {
  readonly kind: 'pi';
  target: string;
  data: string;
}

export type XmlNode = XmlElement | XmlText | XmlCData | XmlComment | XmlPI;

export interface XmlDeclaration {
  readonly version: string;
  readonly encoding?: string;
  readonly standalone?: 'yes' | 'no';
}

export interface XmlDocument {
  readonly kind: 'document';
  decl: XmlDeclaration | null;
  /** Comments / processing instructions appearing before the root element. */
  prolog: ReadonlyArray<XmlComment | XmlPI>;
  root: XmlElement;
  /** Comments / processing instructions appearing after the root element. */
  epilog: ReadonlyArray<XmlComment | XmlPI>;
}

// Constructor helpers — explicit verbs rather than `new` invocations so call
// sites read naturally.

export const qname = (prefix: string, localName: string, namespaceURI: string): QName => ({
  prefix,
  localName,
  namespaceURI,
});

export const attr = (name: QName, value: string): XmlAttr => ({ name, value });

export const elem = (
  name: QName,
  init: {
    attrs?: XmlAttr[];
    prefixDecls?: Map<string, string>;
    children?: XmlNode[];
  } = {},
): XmlElement => ({
  kind: 'element',
  name,
  attrs: init.attrs ?? [],
  prefixDecls: init.prefixDecls ?? new Map(),
  children: init.children ?? [],
});

export const text = (data: string): XmlText => ({ kind: 'text', data });
export const cdata = (data: string): XmlCData => ({ kind: 'cdata', data });
export const comment = (data: string): XmlComment => ({ kind: 'comment', data });
export const pi = (target: string, data: string): XmlPI => ({ kind: 'pi', target, data });

// QName comparison — namespace-first. Two QNames are equal iff their URI and
// localName match. Prefix is irrelevant for equality (it's a serialization
// detail) but is preserved on the values themselves.
export const qnameEquals = (a: QName, b: QName): boolean =>
  a.localName === b.localName && a.namespaceURI === b.namespaceURI;
