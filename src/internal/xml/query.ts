// Query helpers for the XML AST.
//
// These are ergonomic accessors for the read-heavy code in higher layers
// (PresentationML reads slide.xml looking for `p:cSld/p:spTree`, etc.).
// They never mutate the AST; mutation goes through the AST constructors and
// direct field assignment.

import type { QName, XmlAttr, XmlElement, XmlNode, XmlText } from './ast.ts';
import { qnameEquals } from './ast.ts';

export const isElement = (n: XmlNode): n is XmlElement => n.kind === 'element';
export const isText = (n: XmlNode): n is XmlText => n.kind === 'text';

export const childElements = (e: XmlElement): XmlElement[] => e.children.filter(isElement);

export const firstChildElement = (e: XmlElement, name: QName): XmlElement | null => {
  for (const c of e.children) {
    if (isElement(c) && qnameEquals(c.name, name)) return c;
  }
  return null;
};

export const allChildElements = (e: XmlElement, name: QName): XmlElement[] => {
  const out: XmlElement[] = [];
  for (const c of e.children) {
    if (isElement(c) && qnameEquals(c.name, name)) out.push(c);
  }
  return out;
};

export const getAttr = (e: XmlElement, name: QName): XmlAttr | null => {
  for (const a of e.attrs) {
    if (qnameEquals(a.name, name)) return a;
  }
  return null;
};

export const getAttrValue = (e: XmlElement, name: QName): string | null =>
  getAttr(e, name)?.value ?? null;

/**
 * Returns the concatenated text content of an element, walking into nested
 * elements. OOXML uses this pattern for `<a:t>Hello</a:t>` (single text
 * child) and for run sequences where multiple text nodes need merging.
 */
export const textContent = (e: XmlElement): string => {
  let out = '';
  for (const c of e.children) {
    if (c.kind === 'text') out += c.data;
    else if (c.kind === 'cdata') out += c.data;
    else if (c.kind === 'element') out += textContent(c);
  }
  return out;
};

/**
 * Walks the element tree depth-first, yielding every element including the
 * root. The visitor returns `false` to skip descending into the current
 * element's children. Returning nothing (or `true`) descends as usual.
 */
export const walkElements = (root: XmlElement, visit: (e: XmlElement) => boolean | void): void => {
  const stack: XmlElement[] = [root];
  while (stack.length > 0) {
    const node = stack.pop();
    if (!node) continue;
    const descend = visit(node);
    if (descend === false) continue;
    // Push in reverse so visit order is DFS-natural left-to-right.
    for (let i = node.children.length - 1; i >= 0; i--) {
      const c = node.children[i];
      if (c && isElement(c)) stack.push(c);
    }
  }
};
