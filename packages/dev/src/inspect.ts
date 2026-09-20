import { readFile } from 'node:fs/promises';
import {
  loadPresentation,
  getSlides,
  getSlidePartName,
  getSlideTitle,
  getSlideShapes,
  getShapeName,
  getShapeId,
  getShapeKind,
  getShapePlaceholderType,
  getShapePlaceholderIdx,
  getSlideLayouts,
  getSlideLayoutName,
  getSlideLayoutPartName,
  getSlideLayoutType,
} from '@office-kit/pptx';

/** Source references usable directly in Slide, Fill and Remove props. */
export async function inspectTemplate(file: string) {
  const presentation = await loadPresentation(await readFile(file));
  return {
    slides: getSlides(presentation).map((slide, index) => ({
      index,
      part: getSlidePartName(slide),
      title: getSlideTitle(slide),
      shapes: getSlideShapes(slide).map((shape) => ({
        id: getShapeId(shape),
        name: getShapeName(shape),
        kind: getShapeKind(shape),
        placeholder: { type: getShapePlaceholderType(shape), idx: getShapePlaceholderIdx(shape) },
      })),
    })),
    layouts: getSlideLayouts(presentation).map((layout) => ({
      part: getSlideLayoutPartName(layout),
      name: getSlideLayoutName(layout),
      type: getSlideLayoutType(layout),
    })),
  };
}
