// The proofing language of a shape's text (`lang` on `a:rPr` / `a:endParaRPr`,
// ST_TextLanguageID): what Review ▸ Language sets. It is not appearance, so
// it is kept apart from the run-format readers and writers.

import {
  NS,
  type XmlElement,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  qname,
  walkElements,
} from '../../internal/xml/index.ts';
import {
  SHAPE_ELEMENT,
  SHAPE_SNAPSHOT,
  type SlideShapeData,
  type TableCellData,
} from '../_internal-symbols.ts';
import { commitAndRefresh, ensureTxBody } from './_helpers.ts';
import { requireParagraph } from './shape-runs.ts';

const ATTR_LANG = qname('', 'lang', '');
const NAME_RPR = qname('a', 'rPr', NS.dml);
const NAME_TX_BODY = qname('p', 'txBody', NS.pml);
// BCP 47: a 2–8 letter primary subtag, then alphanumeric subtags.
const LANGUAGE_TAG = /^[A-Za-z]{2,8}(?:-[A-Za-z0-9]{1,8})*$/;
// Elements whose CT type takes an optional `a:rPr` as their first child.
const RUN_CONTAINERS = new Set(['r', 'fld', 'br']);

const setLang = (properties: XmlElement, lang: string): void => {
  properties.attrs = [
    ...properties.attrs.filter((a) => !(a.name.namespaceURI === '' && a.name.localName === 'lang')),
    attr(ATTR_LANG, lang),
  ];
};

/** Marks every run in the shape's text (and the end-of-paragraph marks) as `lang`, e.g. `en-US`. */
export const setShapeTextLanguage = (shape: SlideShapeData, lang: string): void => {
  if (!LANGUAGE_TAG.test(lang))
    throw new Error(`setShapeTextLanguage: "${lang}" is not a language tag`);
  const txBody = ensureTxBody(shape);
  walkElements(txBody, (element) => {
    if (element.name.namespaceURI !== NS.dml) return;
    if (element.name.localName === 'endParaRPr') setLang(element, lang);
    else if (RUN_CONTAINERS.has(element.name.localName)) {
      let properties = firstChildElement(element, NAME_RPR);
      if (!properties) {
        properties = elem(NAME_RPR);
        element.children.unshift(properties);
      }
      setLang(properties, lang);
    }
  });
  commitAndRefresh(shape);
};

/** The language of the shape's first run that names one, or `null`. */
export const getShapeTextLanguage = (shape: SlideShapeData): string | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'shape') return null;
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  let found: string | null = null;
  walkElements(txBody, (element) => {
    if (found !== null || element.name.namespaceURI !== NS.dml) return;
    if (element.name.localName === 'rPr' || element.name.localName === 'endParaRPr')
      found = getAttrValue(element, ATTR_LANG);
  });
  return found;
};

/**
 * The literal `lang` of each run, field and line break in a paragraph, in
 * `getShapeParagraphElements` order, or `null` where the element names none.
 * Renderers use it for language-dependent layout such as the character a
 * decimal tab aligns on.
 */
export const getParagraphElementLanguages = (
  target: SlideShapeData | TableCellData,
  paragraphIndex: number,
): ReadonlyArray<string | null> => {
  const languages: (string | null)[] = [];
  for (const child of requireParagraph(target, paragraphIndex).children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.dml) continue;
    if (!RUN_CONTAINERS.has(child.name.localName)) continue;
    const properties = firstChildElement(child, NAME_RPR);
    languages.push(properties ? getAttrValue(properties, ATTR_LANG) : null);
  }
  return languages;
};
