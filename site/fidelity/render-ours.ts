// Render every slide of a .pptx with pptx-kit's own previewer and rasterize it
// to pixels — no browser. The whole SVG→resvg→RGBA pipeline now lives in
// `pptx-kit-preview/node`; the harness only feeds it a deck and a width.

import { getSlides, getSlideSize, type PresentationData } from 'pptx-kit';
import { loadPresentationFile } from 'pptx-kit/node';
import { renderSlideToRgba } from 'pptx-kit-preview/node';
import type { RgbaImage } from './image.ts';

export interface RenderedSlide {
  readonly image: RgbaImage;
  readonly png: Uint8Array;
}

// 16:9 widescreen fallback, matching the renderer's own default.
const DEFAULT_W = 12_192_000;
const DEFAULT_H = 6_858_000;

export interface RenderOptions {
  /** Target raster width in pixels; height follows the slide aspect ratio. */
  readonly width: number;
}

export interface PresentationRender {
  readonly slides: RenderedSlide[];
  readonly pixelWidth: number;
  readonly pixelHeight: number;
}

export const renderPresentation = async (
  pptxPath: string,
  opts: RenderOptions,
): Promise<PresentationRender> => {
  const pres: PresentationData = await loadPresentationFile(pptxPath);
  const size = getSlideSize(pres) ?? { width: DEFAULT_W, height: DEFAULT_H };
  const emuW = size.width as number;
  const emuH = size.height as number;
  const pixelWidth = opts.width;
  const pixelHeight = Math.round((opts.width * emuH) / emuW);

  const slides = getSlides(pres).map(
    (slide): RenderedSlide => renderSlideToRgba(pres, slide, { width: pixelWidth }),
  );

  return { slides, pixelWidth, pixelHeight };
};
