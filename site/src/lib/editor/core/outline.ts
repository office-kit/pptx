import {
  getShapeId,
  getShapePlaceholderType,
  getSlideShapes,
  getShapeKind,
  isShapePlaceholder,
  type SlideData,
} from '@office-kit/pptx';

/** PowerPoint's outline excludes ordinary text boxes and footer placeholders. */
export function outlineShapes(slide: SlideData) {
  return getSlideShapes(slide).flatMap((shape) => {
    if (!isShapePlaceholder(shape) || getShapeKind(shape) !== 'shape') return [];
    const type = getShapePlaceholderType(shape) ?? 'obj';
    if (!['title', 'ctrTitle', 'subTitle', 'body', 'obj'].includes(type)) return [];
    return [{ id: getShapeId(shape), title: type === 'title' || type === 'ctrTitle' }];
  });
}
