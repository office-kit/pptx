// internal/xml — namespace-aware XML AST.
// Allowed imports: none.
// See plan §Architecture — XML.

export type {
  QName,
  XmlAttr,
  XmlCData,
  XmlComment,
  XmlDeclaration,
  XmlDocument,
  XmlElement,
  XmlNode,
  XmlPI,
  XmlText,
} from './ast.ts';
export { attr, cdata, comment, elem, pi, qname, qnameEquals, text } from './ast.ts';
export { NS, SUGGESTED_PREFIX } from './namespaces.ts';
export { XmlParseError, parseFragment, parseXml } from './parse.ts';
export {
  allChildElements,
  childElements,
  firstChildElement,
  getAttr,
  getAttrValue,
  isElement,
  isText,
  textContent,
  walkElements,
} from './query.ts';
export { serializeFragment, serializeXml } from './serialize.ts';
