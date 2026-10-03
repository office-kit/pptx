import {
  addBlankSlide,
  asColor,
  createPresentation,
  inches,
  setSlideBackground,
  setSlideBackgroundGradientFill,
  setSlideSize,
  type SlideMasterBackgroundStyle,
} from '@office-kit/pptx';
import { renderSlideToSvg } from '@office-kit/pptx-preview';

/** Render swatches with the same gradient geometry as the slide canvas. */
export function backgroundStyleSwatches(
  styles: readonly SlideMasterBackgroundStyle[],
): ReadonlyMap<number, string> {
  const swatches = new Map<number, string>();
  if (!styles.length) return swatches;
  const pres = createPresentation();
  setSlideSize(pres, { width: inches(1.5), height: inches(1) });
  const slide = addBlankSlide(pres);
  for (const style of styles) {
    if (style.gradient) {
      setSlideBackgroundGradientFill(slide, {
        ...style.gradient,
        // The reader has already applied theme transforms. Reapplying them to
        // resolved RGB values would tint/shade a second time.
        stops: style.gradient.stops.map((stop) => ({
          offset: stop.offset,
          color: asColor(stop.resolvedColor ?? stop.color) ?? '#000000',
          ...(stop.opacity !== undefined ? { opacity: stop.opacity } : {}),
        })),
      });
    } else if (style.fill.kind === 'solid') {
      setSlideBackground(slide, asColor(style.fill.color) ?? '#FFFFFF', style.fill.opacity);
    } else {
      continue;
    }
    swatches.set(
      style.style,
      `data:image/svg+xml,${encodeURIComponent(renderSlideToSvg(pres, slide))}`,
    );
  }
  return swatches;
}
