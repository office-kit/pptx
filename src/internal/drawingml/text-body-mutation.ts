// Mutating helpers for `a:txBody`. Companion to `text-body.ts`, which is
// read-only.
//
// Why a separate file: the read-only path is hot (called on every slide
// load), the mutation path runs only when the user calls a `set*` method.
// Keeping them separate keeps `text-body.ts` small and clearly side-effect-
// free.

import {
  NS,
  type XmlAttr,
  type XmlElement,
  type XmlNode,
  attr,
  elem,
  firstChildElement,
  qname,
  text,
  walkElements,
} from '../xml/index.ts';

const NAME_P = qname('a', 'p', NS.dml);
const NAME_R = qname('a', 'r', NS.dml);
const NAME_T = qname('a', 't', NS.dml);
const NAME_RPR = qname('a', 'rPr', NS.dml);
const NAME_PPR = qname('a', 'pPr', NS.dml);
const ATTR_XML_SPACE = qname('xml', 'space', NS.xml);

const cloneAttrs = (attrs: ReadonlyArray<XmlAttr>): XmlAttr[] =>
  attrs.map((a) => attr(a.name, a.value));

const cloneElement = (e: XmlElement): XmlElement =>
  elem(e.name, {
    attrs: cloneAttrs(e.attrs),
    prefixDecls: new Map(e.prefixDecls),
    children: e.children.map(cloneNode),
  });

const cloneNode = (n: XmlNode): XmlNode => {
  switch (n.kind) {
    case 'element':
      return cloneElement(n);
    case 'text':
      return { kind: 'text', data: n.data };
    case 'cdata':
      return { kind: 'cdata', data: n.data };
    case 'comment':
      return { kind: 'comment', data: n.data };
    case 'pi':
      return { kind: 'pi', target: n.target, data: n.data };
  }
};

/**
 * Returns the first `a:rPr` reachable on the first run of the first paragraph
 * inside `txBody`. Used to preserve formatting on text replacement: we clone
 * the original rPr into the new run so font, color, size, etc. survive.
 */
const findFirstRunProperties = (txBody: XmlElement): XmlElement | null => {
  const firstP = firstChildElement(txBody, NAME_P);
  if (firstP === null) return null;
  const firstR = firstChildElement(firstP, NAME_R);
  if (firstR === null) return null;
  return firstChildElement(firstR, NAME_RPR);
};

/**
 * Returns the first `a:pPr` from the first paragraph. Used to preserve
 * paragraph-level formatting (alignment, bullet, indent) when replacing.
 */
const findFirstParagraphProperties = (txBody: XmlElement): XmlElement | null => {
  const firstP = firstChildElement(txBody, NAME_P);
  if (firstP === null) return null;
  return firstChildElement(firstP, NAME_PPR);
};

const removeAllParagraphs = (txBody: XmlElement): void => {
  txBody.children = txBody.children.filter(
    (c) => !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'p'),
  );
};

// Token pattern: `{{key}}` where `key` matches any non-brace characters.
// This is the same syntax docxtemplater / Handlebars use, and it survives
// XML escaping cleanly since `{` and `}` are not XML metacharacters.
const TOKEN_PATTERN = /\{\{([^{}]+)\}\}/g;

/**
 * Replaces `{{key}}` tokens in every `a:t` element under `root` with values
 * from `tokens`. Tokens whose key is not in `tokens` are left untouched.
 *
 * Returns the number of substitutions performed. Useful for callers that
 * want to know whether anything matched.
 *
 * Limitation: a token must fit entirely within one `<a:t>` element to
 * match. PowerPoint normally serializes contiguous user text as a single
 * `<a:t>`, so the limitation only bites when a placeholder was edited
 * character-by-character (causing PowerPoint to split runs). For those
 * cases, fall back to `setText()`.
 */
export const replaceTokensInTree = (root: XmlElement, tokens: Record<string, string>): number => {
  let count = 0;
  walkElements(root, (el) => {
    if (el.name.namespaceURI !== NAME_T.namespaceURI) return;
    if (el.name.localName !== 't') return;
    const child = el.children[0];
    if (!child || child.kind !== 'text') return;
    const before = child.data;
    let didMatch = false;
    const after = before.replace(TOKEN_PATTERN, (match, key: string) => {
      if (Object.hasOwn(tokens, key)) {
        didMatch = true;
        return tokens[key] ?? '';
      }
      return match;
    });
    if (didMatch) {
      child.data = after;
      count++;
    }
  });
  return count;
};

/**
 * Replaces the entire visible text of a `txBody` element with `value`.
 * Newlines in `value` start a new paragraph.
 *
 * The first existing `a:rPr` (run properties) and `a:pPr` (paragraph
 * properties) are cloned into every new paragraph so that font, color, size,
 * alignment, and bullet style survive the replacement. If the source had
 * mixed formatting per run, that gets collapsed into the first-run formatting
 * across the entire new text — matching what PowerPoint does when you select
 * all and type.
 *
 * The `bodyPr` and `lstStyle` children (if any) are preserved untouched.
 */
export const setTextBody = (txBody: XmlElement, value: string): void => {
  const rPrTemplate = findFirstRunProperties(txBody);
  const pPrTemplate = findFirstParagraphProperties(txBody);

  removeAllParagraphs(txBody);

  const lines = value.split('\n');
  for (const line of lines) {
    const t = elem(NAME_T, {
      attrs:
        line.length > 0 && (line.startsWith(' ') || line.endsWith(' ') || line.includes('\t'))
          ? [attr(ATTR_XML_SPACE, 'preserve')]
          : [],
      children: line.length > 0 ? [text(line)] : [],
    });
    const r = elem(NAME_R, {
      children: rPrTemplate !== null ? [cloneElement(rPrTemplate), t] : [t],
    });
    const p = elem(NAME_P, {
      children: pPrTemplate !== null ? [cloneElement(pPrTemplate), r] : [r],
    });
    txBody.children.push(p);
  }
};
