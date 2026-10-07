import {
  getPresentationFonts,
  getShapePlaceholderType,
  isShapePlaceholder,
  isShapeTextBox,
  type PresentationData,
  type SlideShapeData,
} from '@office-kit/pptx';

export const DEFAULT_BODY_PT = 18;
const DEFAULT_TITLE_PT = 44;
const DEFAULT_SUBTITLE_PT = 32;
const DEFAULT_FOOTER_PT = 12;
const DEFAULT_FONT = `Calibri, "Helvetica Neue", Arial, sans-serif`;

/** Match the preview's fallback metrics when no authored or inherited font is available. */
export function defaultTextMetrics(pres: PresentationData, shape: SlideShapeData) {
  const placeholder = getShapePlaceholderType(shape);
  const size =
    placeholder === 'title' || placeholder === 'ctrTitle'
      ? DEFAULT_TITLE_PT
      : placeholder === 'subTitle'
        ? DEFAULT_SUBTITLE_PT
        : placeholder === 'ftr' || placeholder === 'dt' || placeholder === 'sldNum'
          ? DEFAULT_FOOTER_PT
          : DEFAULT_BODY_PT;
  const fonts = getPresentationFonts(pres);
  const face =
    placeholder === 'title' || placeholder === 'ctrTitle' ? fonts?.majorLatin : fonts?.minorLatin;
  return { size, family: face ? `${JSON.stringify(face)}, ${DEFAULT_FONT}` : DEFAULT_FONT };
}

/** Match the preview when a shape has no authored or inherited body/paragraph alignment. */
export function shapeTextDefaults(shape: SlideShapeData): {
  align: 'left' | 'center';
  anchor: 'top' | 'center';
} {
  return isShapePlaceholder(shape) || isShapeTextBox(shape)
    ? { align: 'left', anchor: 'top' }
    : { align: 'center', anchor: 'center' };
}
