import { readGroupChildren } from '../../internal/presentationml/index.ts';
import { NS, type XmlElement, firstChildElement, qname } from '../../internal/xml/index.ts';
import {
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  SLIDE_SHAPES,
  type SlideShapeData,
} from '../_internal-symbols.ts';

const parentMaps = new WeakMap<object, WeakMap<object, XmlElement>>();

// Structural edits rebuild SLIDE_SHAPES, invalidating this parent index.
// Index all descendants once so resolving every child does not scan the slide repeatedly.
const getParentMap = (shape: SlideShapeData): WeakMap<object, XmlElement> => {
  const slide = shape[SHAPE_SLIDE];
  const shapeList = slide[SLIDE_SHAPES];
  let byShape = parentMaps.get(shapeList);
  if (!byShape) {
    byShape = new WeakMap();
    const walk = (group: XmlElement): void => {
      for (const child of readGroupChildren(group)) {
        byShape!.set(child.element, group);
        if (child.kind === 'group') walk(child.element);
      }
    };
    for (const candidate of shapeList) {
      if (candidate[SHAPE_SNAPSHOT].kind === 'group') walk(candidate[SHAPE_ELEMENT]);
    }
    parentMaps.set(shapeList, byShape);
  }
  return byShape;
};

export const containingGroupFillElement = (shape: SlideShapeData): XmlElement | null => {
  const parents = getParentMap(shape);
  let group = parents.get(shape[SHAPE_ELEMENT]);
  while (group) {
    const props = firstChildElement(group, qname('p', 'grpSpPr', NS.pml));
    const fill = props?.children.find(
      (node): node is XmlElement =>
        node.kind === 'element' &&
        node.name.namespaceURI === NS.dml &&
        ['noFill', 'solidFill', 'gradFill', 'pattFill', 'blipFill', 'grpFill'].includes(
          node.name.localName,
        ),
    );
    if (fill && fill.name.localName !== 'grpFill') return fill;
    group = parents.get(group);
  }
  return null;
};
