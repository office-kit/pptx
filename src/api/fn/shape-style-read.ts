// Shape fill references use the DrawingML theme format scheme.

import { getEffectiveColorMap } from './color-map.ts';
import { resolveDrawingColor } from './shape-color.ts';
import {
  getPresentationTheme,
  themeFromPackage,
  themeRootFromPackage,
  type PresentationTheme,
} from './theme.ts';
import { getSlideLayout } from './shape-slide-read.ts';
import { resolveTarget } from '../../internal/opc/index.ts';
import { REL_TYPES } from '../../internal/presentationml/index.ts';
import {
  NS,
  cloneElement,
  elem,
  type XmlElement,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  type PresentationData,
  LAYOUT_PART_NAME,
  type SlideShapeData,
} from '../_internal-symbols.ts';

export type ShapeStyleFill =
  | { readonly kind: 'solid'; readonly color: string; readonly colorElement: XmlElement }
  | { readonly kind: 'gradient' }
  | { readonly kind: 'pattern' }
  | { readonly kind: 'image' }
  | { readonly kind: 'none' };

const NAME_STYLE = qname('p', 'style', NS.pml);
const NAME_FILL_REF = qname('a', 'fillRef', NS.dml);
const NAME_LN_REF = qname('a', 'lnRef', NS.dml);
const NAME_FMT_SCHEME = qname('a', 'fmtScheme', NS.dml);

type ShapeStyleTheme = {
  readonly root: XmlElement | null;
  readonly theme: PresentationTheme | null;
};

/**
 * Returns the theme attached to the shape's owning slide master. A package
 * can contain multiple slide masters with different themes; using the first
 * theme part would make a shape on a later master render with the wrong fill.
 */
export const getShapeStyleTheme = (
  pres: PresentationData,
  shape: SlideShapeData,
): ShapeStyleTheme => {
  const pkg = pres[INTERNAL_PACKAGE];
  const layout = getSlideLayout(shape[SHAPE_SLIDE]);
  if (layout) {
    const layoutName = layout[LAYOUT_PART_NAME];
    const masterRel = pkg
      .getRels(layoutName)
      ?.items.find((item) => item.type === REL_TYPES.slideMaster && item.targetMode !== 'External');
    if (masterRel) {
      const masterName = resolveTarget(layoutName, masterRel.target);
      return {
        root: themeRootFromPackage(pkg, masterName),
        theme: themeFromPackage(pkg, masterName),
      };
    }
  }
  return {
    root: themeRootFromPackage(pkg),
    theme: getPresentationTheme(pres),
  };
};

const replaceStyleColorPlaceholder = (resolved: XmlElement, reference: XmlElement): void => {
  const referenceColor = reference.children.find(
    (child): child is XmlElement => child.kind === 'element' && child.name.namespaceURI === NS.dml,
  );
  if (!referenceColor) return;
  const replace = (parent: XmlElement): void => {
    parent.children = parent.children.map((child) => {
      if (child.kind !== 'element') return child;
      if (
        child.name.localName === 'schemeClr' &&
        getAttrValue(child, qname('', 'val', '')) === 'phClr'
      ) {
        const replacement = cloneElement(referenceColor);
        replacement.children.push(
          ...child.children.map((item) =>
            item.kind === 'element' ? cloneElement(item) : { ...item },
          ),
        );
        return replacement;
      }
      replace(child);
      return child;
    });
  };
  replace(resolved);
};

const readShapeStyleReferenceElement = (
  pres: PresentationData,
  shape: SlideShapeData,
  referenceName: XmlElement['name'],
  listLocalName: 'fillStyleLst' | 'lnStyleLst',
): XmlElement | null => {
  const style = firstChildElement(shape[SHAPE_ELEMENT], NAME_STYLE);
  const reference = style ? firstChildElement(style, referenceName) : null;
  if (!reference) return null;

  const index = Number.parseInt(getAttrValue(reference, qname('', 'idx', '')) ?? '', 10);
  if (!Number.isInteger(index)) return null;
  if (index === 0 || (index === 1000 && listLocalName === 'fillStyleLst'))
    return elem(qname('a', 'noFill', NS.dml));
  if (index < 1) return null;

  const themeRoot = getShapeStyleTheme(pres, shape).root;
  const elements = themeRoot
    ? firstChildElement(themeRoot, qname('a', 'themeElements', NS.dml))
    : null;
  const scheme = elements ? firstChildElement(elements, NAME_FMT_SCHEME) : null;
  const list = scheme
    ? firstChildElement(
        scheme,
        qname(
          'a',
          index >= 1001 && listLocalName === 'fillStyleLst' ? 'bgFillStyleLst' : listLocalName,
          NS.dml,
        ),
      )
    : null;
  const styles = list?.children.filter(
    (child): child is XmlElement => child.kind === 'element' && child.name.namespaceURI === NS.dml,
  );
  const backgroundFill = listLocalName === 'fillStyleLst' && index >= 1001;
  const selected = styles?.[backgroundFill ? index - 1001 : index - 1];
  if (!selected) return null;

  const resolved = cloneElement(selected);
  replaceStyleColorPlaceholder(resolved, reference);

  return resolved;
};

/** Resolves a shape's fillRef against the presentation theme format scheme. */
export const readShapeStyleFillElement = (
  pres: PresentationData,
  shape: SlideShapeData,
): XmlElement | null => readShapeStyleReferenceElement(pres, shape, NAME_FILL_REF, 'fillStyleLst');

/** Resolves a shape's lnRef against the owning master's line style list. */
export const readShapeStyleLineElement = (
  pres: PresentationData,
  shape: SlideShapeData,
): XmlElement | null => readShapeStyleReferenceElement(pres, shape, NAME_LN_REF, 'lnStyleLst');

/** Reads the resolved fill choice from the shape's own style reference. */
export const readShapeStyleFill = (
  pres: PresentationData,
  shape: SlideShapeData,
): ShapeStyleFill | null => {
  const resolved = readShapeStyleFillElement(pres, shape);
  if (!resolved) return null;

  switch (resolved.name.localName) {
    case 'solidFill': {
      const colorElement = resolved.children.find(
        (child): child is XmlElement =>
          child.kind === 'element' && child.name.namespaceURI === NS.dml,
      );
      if (!colorElement) return { kind: 'solid', color: '', colorElement: resolved };
      const theme = getShapeStyleTheme(pres, shape).theme;
      const color = resolveDrawingColor(
        colorElement,
        theme,
        getEffectiveColorMap(shape[SHAPE_SLIDE]),
      );
      return {
        kind: 'solid',
        color: color ?? '',
        colorElement,
      };
    }
    case 'gradFill':
      return { kind: 'gradient' };
    case 'pattFill':
      return { kind: 'pattern' };
    case 'blipFill':
      return { kind: 'image' };
    case 'noFill':
      return { kind: 'none' };
    default:
      return null;
  }
};
