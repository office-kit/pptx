import {
  getSlideBackground,
  getSlideBackgroundGradientFill,
  getSlideBackgroundPatternFill,
  getSlideLayout,
  getSlideLayoutBackground,
  getSlideLayoutBackgroundGradientFill,
  getSlideLayoutBackgroundPatternFill,
  getSlideMasterBackground,
  getSlideMasterBackgroundGradientFill,
  getSlideMasterBackgroundPatternFill,
  type PresentationData,
  type SlideData,
} from '@office-kit/pptx';

/** Resolve the visible fill before editing; an explicit override blocks inherited fills. */
export function readSlideBackground(
  pres: PresentationData,
  slide: SlideData,
  options: { preserveTheme?: boolean } = {},
) {
  const fill = getSlideBackground(slide);
  if (fill.kind !== 'inherit')
    return {
      fill,
      gradient: getSlideBackgroundGradientFill(slide),
      pattern: getSlideBackgroundPatternFill(pres, slide, options),
    };
  const layout = getSlideLayout(slide);
  if (!layout) return { fill, gradient: null, pattern: null };
  const layoutFill = getSlideLayoutBackground(layout);
  if (layoutFill.kind !== 'inherit')
    return {
      fill: layoutFill,
      gradient: getSlideLayoutBackgroundGradientFill(layout),
      pattern: getSlideLayoutBackgroundPatternFill(pres, layout, options),
    };
  return {
    fill: getSlideMasterBackground(pres, layout),
    gradient: getSlideMasterBackgroundGradientFill(pres, layout),
    pattern: getSlideMasterBackgroundPatternFill(pres, layout, options),
  };
}
