import { isShapePlaceholder, isShapeTextBox, type SlideShapeData } from '@office-kit/pptx';

/** Match the preview when a shape has no authored or inherited body/paragraph alignment. */
export function shapeTextDefaults(shape: SlideShapeData): {
  align: 'left' | 'center';
  anchor: 'top' | 'center';
} {
  return isShapePlaceholder(shape) || isShapeTextBox(shape)
    ? { align: 'left', anchor: 'top' }
    : { align: 'center', anchor: 'center' };
}
