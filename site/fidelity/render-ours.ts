// Render every slide of a .pptx with pptx-kit's own previewer
// (`renderSlideSvg`) and rasterize it to pixels with resvg — no browser.
//
// Text is laid out as pure SVG (textLayout: 'svg') using a fontkit measurer so
// resvg can rasterize it. The measurer, resvg's `fontFiles`, and the
// LibreOffice ground-truth host all use the SAME bundled substitute fonts, so
// the engine's wrap/position math agrees with the painted pixels.

import { Resvg } from '@resvg/resvg-js';
import { getSlides, getSlideSize, type PresentationData } from 'pptx-kit';
import { loadPresentationFile } from 'pptx-kit/node';
import { renderSlideSvg } from '../src/lib/playground/render-slide.ts';
import { buildFontkitMeasurer, FONT_FILES } from './measure.ts';
import { SANS, SERIF, MONO } from '../src/lib/playground/text-layout.ts';
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

  // Build the measurer once (loads + verifies every face), reuse per slide.
  const measureText = buildFontkitMeasurer();

  const slides = getSlides(pres).map((slide): RenderedSlide => {
    const svg = renderSlideSvg(pres, slide, { measureText, textLayout: 'svg' });
    const resvg = new Resvg(svg, {
      fitTo: { mode: 'width', value: pixelWidth },
      font: {
        loadSystemFonts: false, // deterministic; only our bundled fonts
        fontFiles: FONT_FILES,
        defaultFontFamily: SANS,
        sansSerifFamily: SANS,
        serifFamily: SERIF,
        monospaceFamily: MONO,
      },
    });
    const rendered = resvg.render();
    const image: RgbaImage = {
      width: rendered.width,
      height: rendered.height,
      data: new Uint8Array(rendered.pixels),
    };
    return { image, png: new Uint8Array(rendered.asPng()) };
  });

  return { slides, pixelWidth, pixelHeight };
};
