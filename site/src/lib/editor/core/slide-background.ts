import {
  getSlideBackground,
  getSlideBackgroundGradientFill,
  getSlideLayout,
  getSlideLayoutBackground,
  getSlideLayoutBackgroundGradientFill,
  getSlideMasterBackground,
  getSlideMasterBackgroundGradientFill,
  type PresentationData,
  type SlideData,
} from '@office-kit/pptx';

/** Resolve the visible fill before editing; a non-gradient override blocks inherited gradients. */
export function readSlideBackground(pres: PresentationData, slide: SlideData) {
  const fill = getSlideBackground(slide);
  if (fill.kind !== 'inherit') return { fill, gradient: getSlideBackgroundGradientFill(slide) };
  const layout = getSlideLayout(slide);
  if (!layout) return { fill, gradient: null };
  const layoutFill = getSlideLayoutBackground(layout);
  if (layoutFill.kind !== 'inherit')
    return { fill: layoutFill, gradient: getSlideLayoutBackgroundGradientFill(layout) };
  return {
    fill: getSlideMasterBackground(pres, layout),
    gradient: getSlideMasterBackgroundGradientFill(pres, layout),
  };
}
