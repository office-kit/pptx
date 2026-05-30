// Render every slide of a .pptx with pptx-kit's own previewer
// (`renderSlideSvg`) and rasterize it to pixels with resvg — no browser.
//
// Caveat that this harness is designed to surface: the current previewer lays
// text out in <foreignObject>, which resvg (like any browser-free SVG
// rasterizer) cannot render. So text is absent from these rasters today. That
// is intentional — the resulting SSIM gap quantifies exactly the "move text to
// pure SVG" work the roadmap puts in Phase 1.

import { Resvg } from '@resvg/resvg-js';
import { getSlides, getSlideSize, type PresentationData } from 'pptx-kit';
import { loadPresentationFile } from 'pptx-kit/node';
import { renderSlideSvg } from '../src/lib/playground/render-slide.ts';
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

  const slides = getSlides(pres).map((slide): RenderedSlide => {
    const svg = renderSlideSvg(pres, slide);
    const resvg = new Resvg(svg, {
      fitTo: { mode: 'width', value: pixelWidth },
      // System-font loading is slow and non-deterministic; the foreignObject
      // text path does not reach resvg anyway.
      font: { loadSystemFonts: false },
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
