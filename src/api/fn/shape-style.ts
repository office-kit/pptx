// Shape style references (the OOXML <p:style> element).

import type { Color, ColorTransform } from '../../internal/drawingml/index.ts';
import { buildColorElement } from '../../internal/drawingml/index.ts';
import { buildColorTransforms } from '../../internal/drawingml/color-transforms.ts';
import { NS, attr, elem, qname } from '../../internal/xml/index.ts';
import { SHAPE_ELEMENT, type SlideShapeData } from '../_internal-symbols.ts';
import { commitAndRefresh, requireSpPr } from './_helpers.ts';

/** One DrawingML style-list reference used by a shape quick style. */
export interface ShapeStyleReference {
  /** 0 clears the reference; positive values index the theme format scheme. */
  readonly idx: number;
  /** Theme slot or sRGB color replacing the style list's `phClr`. */
  readonly color?: Color;
  /** Ordered DrawingML color transforms (for example the reference desktop app's shade). */
  readonly colorTransforms?: readonly ColorTransform[];
}

export interface ShapeStyleFontReference {
  readonly idx: 'major' | 'minor' | 'none';
  readonly color?: Color;
  readonly colorTransforms?: readonly ColorTransform[];
}

/** Complete OOXML shape style. Direct text runs and geometry are untouched. */
export interface ShapeStyleOptions {
  readonly line: ShapeStyleReference;
  readonly fill: ShapeStyleReference;
  readonly effect: ShapeStyleReference;
  readonly font: ShapeStyleFontReference;
}

const STYLE = qname('p', 'style', NS.pml);
const SPPR_PAINTS = new Set(['noFill', 'solidFill', 'gradFill', 'pattFill', 'blipFill', 'grpFill']);

const colorChild = (
  color: Color | undefined,
  transforms: readonly ColorTransform[] | undefined,
) => {
  if (color === undefined) return undefined;
  const result = buildColorElement(color);
  if (transforms?.length) result.children.push(...buildColorTransforms(transforms));
  return result;
};

const referenceElement = (
  localName: 'lnRef' | 'fillRef' | 'effectRef' | 'fontRef',
  reference: ShapeStyleReference | ShapeStyleFontReference,
) => {
  const idx = reference.idx;
  if (localName === 'fontRef' && idx !== 'major' && idx !== 'minor' && idx !== 'none')
    throw new RangeError(`setShapeStyle: invalid fontRef index ${String(idx)}`);
  if (
    localName !== 'fontRef' &&
    (typeof idx !== 'number' || !Number.isInteger(idx) || idx < 0 || idx > 0xffffffff)
  )
    throw new RangeError(`setShapeStyle: invalid ${localName} index ${idx}`);
  const children = colorChild(reference.color, reference.colorTransforms);
  return elem(qname('a', localName, NS.dml), {
    attrs: [attr(qname('', 'idx', ''), String(idx))],
    ...(children ? { children: [children] } : {}),
  });
};

/** Apply a complete native OOXML shape style while preserving geometry/text. */
export const setShapeStyle = (shape: SlideShapeData, options: ShapeStyleOptions): void => {
  const element = shape[SHAPE_ELEMENT];
  // Construct and validate every reference before touching the document. This
  // keeps a malformed command atomic when dispatched inside an editor action.
  const style = elem(STYLE, {
    children: [
      referenceElement('lnRef', options.line),
      referenceElement('fillRef', options.fill),
      referenceElement('effectRef', options.effect),
      referenceElement('fontRef', options.font),
    ],
  });
  const spPr = requireSpPr(shape);
  // A slide-background fill is stored on the shape itself rather than in
  // spPr. Once a native style is selected it must not continue to override
  // the style's fill reference.
  element.attrs = element.attrs.filter(
    (attribute) =>
      !(attribute.name.namespaceURI === '' && attribute.name.localName === 'useBgFill'),
  );
  spPr.children = spPr.children.filter(
    (child) =>
      child.kind !== 'element' ||
      !(
        child.name.namespaceURI === NS.dml &&
        (child.name.localName === 'ln' ||
          child.name.localName === 'effectLst' ||
          child.name.localName === 'effectDag' ||
          SPPR_PAINTS.has(child.name.localName))
      ),
  );
  element.children = element.children.filter(
    (child) =>
      !(
        child.kind === 'element' &&
        child.name.namespaceURI === NS.pml &&
        child.name.localName === 'style'
      ),
  );
  const spPrIndex = element.children.findIndex(
    (child) =>
      child.kind === 'element' &&
      child.name.namespaceURI === NS.pml &&
      child.name.localName === 'spPr',
  );
  element.children.splice(spPrIndex < 0 ? element.children.length : spPrIndex + 1, 0, style);
  commitAndRefresh(shape);
};
