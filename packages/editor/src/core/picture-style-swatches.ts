import {
  addBlankSlide,
  addSlideImage,
  BUILTIN_PICTURE_STYLES,
  createPresentation,
  inches,
  setShapePictureStyle,
  setSlideBackground,
  setSlideSize,
  type BuiltinPictureStyleName,
} from '@office-kit/pptx';
import { renderSlideToSvg } from '@office-kit/pptx-preview';

// A 96 × 64 landscape (sky over a green hill), the kind of stand-in picture
// the reference desktop app's gallery shows. Generated, not taken from the reference desktop app.
const SAMPLE_PNG =
  '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';

// The swatches depend on nothing but the style list, so they are drawn once.
let cache: ReadonlyMap<BuiltinPictureStyleName, string> | null = null;

/**
 * Each built-in picture style applied to a sample picture, drawn by the
 * preview renderer as an SVG data URL, in gallery order.
 */
export function pictureStyleSwatches(): ReadonlyMap<BuiltinPictureStyleName, string> {
  if (cache) return cache;
  const bytes = Uint8Array.from(atob(SAMPLE_PNG), (char) => char.charCodeAt(0));
  const pres = createPresentation();
  // Room around the picture for frames, shadows and reflections.
  setSlideSize(pres, { width: inches(1.1), height: inches(1) });
  const slide = addBlankSlide(pres);
  // A light gray tile keeps white frames visible.
  setSlideBackground(slide, '#EDEDED');
  const picture = addSlideImage(slide, bytes, {
    x: inches(0.18),
    y: inches(0.16),
    w: inches(0.74),
    h: inches(0.5),
  });
  const images = new Map<BuiltinPictureStyleName, string>();
  for (const style of BUILTIN_PICTURE_STYLES) {
    setShapePictureStyle(picture, style);
    images.set(
      style,
      `data:image/svg+xml,${encodeURIComponent(renderSlideToSvg(pres, slide, { textLayout: 'svg' }))}`,
    );
  }
  cache = images;
  return images;
}
