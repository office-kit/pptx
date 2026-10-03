// Reads the character defaults carried by a shape's DrawingML style reference.

import { getEffectiveColorMap } from './color-map.ts';
import { resolveDrawingColor } from './shape-color.ts';
import { getShapeStyleTheme } from './shape-style-read.ts';
import type { ReadTextFormat } from '../../internal/drawingml/index.ts';
import {
  NS,
  type XmlElement,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';
import {
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  type PresentationData,
  type SlideShapeData,
} from '../_internal-symbols.ts';

const NAME_STYLE = qname('p', 'style', NS.pml);
const NAME_FONT_REF = qname('a', 'fontRef', NS.dml);

/**
 * Reads `<p:style><a:fontRef>` as the shape-level text fallback.
 *
 * PowerPoint's shape Quick Styles use `fontRef` for both the theme font
 * family and the default text color. Mac PowerPoint applies these before
 * inherited placeholder/master defaults, but after direct character and
 * paragraph formatting.
 */
export const readShapeStyleFontFormat = (
  pres: PresentationData,
  shape: SlideShapeData,
): Partial<ReadTextFormat> => {
  const style = firstChildElement(shape[SHAPE_ELEMENT], NAME_STYLE);
  const fontRef = style ? firstChildElement(style, NAME_FONT_REF) : null;
  if (!fontRef) return {};

  const result: Partial<ReadTextFormat> = {};
  const idx = getAttrValue(fontRef, qname('', 'idx', ''));
  if (idx === 'major' || idx === 'minor') {
    const prefix = idx === 'major' ? '+mj' : '+mn';
    result.font = `${prefix}-lt`;
    result.fontEastAsian = `${prefix}-ea`;
    result.fontComplexScript = `${prefix}-cs`;
  }

  const color = fontRef.children.find(
    (child): child is XmlElement => child.kind === 'element' && child.name.namespaceURI === NS.dml,
  );
  if (color) {
    const { theme } = getShapeStyleTheme(pres, shape);
    const resolved = resolveDrawingColor(color, theme, getEffectiveColorMap(shape[SHAPE_SLIDE]));
    if (resolved !== null) result.color = resolved;
  }
  return result;
};
