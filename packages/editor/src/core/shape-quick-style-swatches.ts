import {
  createPresentation,
  getSlideShapes,
  importSlide,
  setShapeHidden,
  setSlideSize,
  setSlideBackground,
  setSlideBackgroundGraphicsHidden,
  addSlideShape,
  setShapeTextAnchor,
  setParagraphAlignment,
  inches,
  type SlideData,
} from '@office-kit/pptx';
import { renderSlideToSvg } from '@office-kit/pptx-preview';
import {
  applyShapeQuickStyle,
  quickStyleColors,
  themeQuickStyles,
  presetQuickStyles,
} from './shape-quick-styles.ts';

/** Use the canvas renderer and the source slide's owning theme for gallery paint. */
export function shapeQuickStyleSwatches(source: SlideData): ReadonlyMap<string, string> {
  const pres = createPresentation();
  // Keep the imported slide itself as the swatch canvas. This preserves a
  // source slide's clrMapOvr, which would be lost when drawing on a fresh
  // blank slide attached only to the imported layout.
  const slide = importSlide(pres, source);
  setShapeHidden(getSlideShapes(slide), true);
  setSlideSize(pres, { width: inches(2), height: inches(1.5) });
  setSlideBackground(slide, '#FFFFFF');
  setSlideBackgroundGraphicsHidden(slide, true);
  const shape = addSlideShape(slide, {
    preset: 'roundRect',
    x: inches(0.3),
    y: inches(0.3),
    w: inches(1.4),
    h: inches(0.9),
    text: 'Abc',
  });
  setShapeTextAnchor(shape, 'center');
  setParagraphAlignment(shape, 0, 'center');
  const images = new Map<string, string>();
  for (const style of [...themeQuickStyles, ...presetQuickStyles]) {
    for (const color of quickStyleColors) {
      applyShapeQuickStyle(shape, style, color);
      images.set(
        `${style}:${color}`,
        `data:image/svg+xml,${encodeURIComponent(renderSlideToSvg(pres, slide, { textLayout: 'svg' }))}`,
      );
    }
  }
  return images;
}
