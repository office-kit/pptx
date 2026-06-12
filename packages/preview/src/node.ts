// `pptx-kit-preview/node` — Node entry: rasterize a slide to a PNG or raw
// RGBA image, with no browser binary.
//
// Pipeline: `renderSlideToSvg(..., { textLayout: 'svg', measureText })` lays
// text out as pure `<text>`, then resvg paints the SVG to pixels. The fontkit
// measurer, resvg's `fontFiles`, and the SVG's family names all reference the
// SAME bundled fonts, so the wrap/positioning math agrees with the rasterized
// glyphs.

import { Resvg } from '@resvg/resvg-js';
import { getSlideSize, type PresentationData, type SlideData } from 'pptx-kit';
import { renderSlideSvg } from './render-slide.ts';
import { MONO, SANS, SERIF, type TextMeasurer } from './text-layout.ts';
import { buildFontkitMeasurer, FONT_FILES } from './measure.ts';

// Re-export the browser-safe surface so a Node consumer gets everything from
// one import.
export * from './index.ts';

// Advanced: pre-build / share a fontkit measurer, or locate the bundled fonts.
export { buildFontkitMeasurer, FONT_FILES, FONT_DIR } from './measure.ts';

/** Raw, un-encoded raster: RGBA bytes in row-major order. */
export interface RgbaImage {
  readonly width: number;
  readonly height: number;
  /** `width * height * 4` bytes, R,G,B,A per pixel. */
  readonly data: Uint8Array;
}

export interface RenderImageOptions {
  /**
   * Target raster width in pixels. Height follows the slide's aspect ratio.
   * Defaults to 1280 (matches the fidelity harness baseline).
   */
  readonly width?: number;
  /**
   * Text measurer used for wrap/positioning. Defaults to a shared fontkit
   * measurer over the bundled fonts. Pass {@link buildFontkitMeasurer}'s
   * result to control its lifetime, or a custom measurer to swap fonts.
   */
  readonly measureText?: TextMeasurer;
}

const DEFAULT_WIDTH = 1280;

// The fontkit measurer reads + verifies every bundled face on construction.
// Build it once and reuse across calls; do it lazily so importing this module
// (e.g. just for the types) doesn't touch the filesystem.
let sharedMeasurer: TextMeasurer | undefined;
const getSharedMeasurer = (): TextMeasurer => (sharedMeasurer ??= buildFontkitMeasurer());

const rasterize = (pres: PresentationData, slide: SlideData, opts: RenderImageOptions) => {
  // `getSlideSize` is only consulted to fail fast on a malformed deck; the SVG
  // carries its own viewBox and resvg fits to the requested width.
  getSlideSize(pres);
  const width = opts.width ?? DEFAULT_WIDTH;
  const measureText = opts.measureText ?? getSharedMeasurer();
  const svg = renderSlideSvg(pres, slide, { measureText, textLayout: 'svg' });
  const resvg = new Resvg(svg, {
    fitTo: { mode: 'width', value: width },
    font: {
      loadSystemFonts: false, // deterministic: only the bundled fonts
      fontFiles: FONT_FILES,
      defaultFontFamily: SANS,
      sansSerifFamily: SANS,
      serifFamily: SERIF,
      monospaceFamily: MONO,
    },
  });
  return resvg.render();
};

/** Render one slide to PNG-encoded bytes. */
export const renderSlideToImage = (
  pres: PresentationData,
  slide: SlideData,
  opts: RenderImageOptions = {},
): Uint8Array => new Uint8Array(rasterize(pres, slide, opts).asPng());

/**
 * Render one slide to raw RGBA pixels plus the PNG encoding of the same frame.
 * Both come from a single rasterization, so callers needing pixel access (SSIM,
 * diffing) and a saveable file don't pay to render twice.
 */
export const renderSlideToRgba = (
  pres: PresentationData,
  slide: SlideData,
  opts: RenderImageOptions = {},
): { readonly image: RgbaImage; readonly png: Uint8Array } => {
  const rendered = rasterize(pres, slide, opts);
  return {
    image: {
      width: rendered.width,
      height: rendered.height,
      data: new Uint8Array(rendered.pixels),
    },
    png: new Uint8Array(rendered.asPng()),
  };
};
