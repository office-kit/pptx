// `@office-kit/pptx-preview/node` — Node entry: rasterize a slide to a PNG or raw
// RGBA image, with no browser binary.
//
// Pipeline: `renderSlideToSvg(..., { textLayout: 'svg', measureText })` lays
// text out as pure `<text>`, then resvg paints the SVG to pixels. The fontkit
// measurer, resvg's `fontFiles`, and the SVG's family names all reference the
// SAME bundled fonts, so the wrap/positioning math agrees with the rasterized
// glyphs.

import { Resvg } from '@resvg/resvg-js';
import * as fontkit from 'fontkit';
import { existsSync, readFileSync } from 'node:fs';
import { getSlideSize, type PresentationData, type SlideData } from '@office-kit/pptx';
import { renderSlideSvg } from './render-slide.ts';
import { MONO, SANS, SERIF, type RenderSlideOptions, type TextMeasurer } from './text-layout.ts';
import { buildFontkitMeasurer, FONT_FILES } from './measure.ts';

// Re-export the browser-safe surface so a Node consumer gets everything from
// one import.
export * from './index.ts';

// Advanced: pre-build / share a fontkit measurer, or locate the bundled fonts.
export { buildFontkitMeasurer, FONT_FILES, FONT_DIR } from './measure.ts';
export type { FontkitMeasurerOptions, RegisteredFont } from './measure.ts';

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
  /**
   * Extra font files (ttf / otf / ttc paths) loaded into the rasterizer in
   * addition to the bundled Latin faces. The bundled set has no CJK
   * coverage, so decks with Japanese / Chinese / Korean text render as
   * missing-glyph boxes unless the caller supplies a face here (pair it with
   * a {@link buildFontkitMeasurer} `fonts` registration so wrap math uses the
   * same metrics). East Asian text is drawn in the family of the first face
   * here that covers kana, in the weight closest to the run's, so supply that
   * family's bold face too for bold Japanese text.
   */
  readonly fontFiles?: ReadonlyArray<string>;
  /**
   * Also let the rasterizer discover OS-installed fonts as a fallback for
   * families the bundled + `fontFiles` faces don't cover. Defaults to
   * `false`: system fonts vary per machine, so the default keeps output
   * deterministic.
   */
  readonly loadSystemFonts?: boolean;
  /**
   * Maps an authored font name to the family drawn and measured, as
   * {@link RenderSlideOptions.resolveFamily}. Defaults to `substituteFamily`.
   * Return a family loaded through `fontFiles` (and registered with the
   * measurer) to draw a deck's own font instead of a bundled substitute.
   * East Asian characters still take the family of the first `fontFiles`
   * face that covers kana; this only decides the rest of the run.
   */
  readonly resolveFamily?: RenderSlideOptions['resolveFamily'];
}

const DEFAULT_WIDTH = 1280;

// The fontkit measurer reads + verifies every bundled face on construction.
// Build it once and reuse across calls; do it lazily so importing this module
// (e.g. just for the types) doesn't touch the filesystem.
let sharedMeasurer: TextMeasurer | undefined;
const getSharedMeasurer = (): TextMeasurer => (sharedMeasurer ??= buildFontkitMeasurer());

// Hiragana あ: a face that has it is a Japanese (or pan-CJK) face.
const KANA_PROBE = 0x3042;

// Reading a face's names means parsing the file, so remember the answer per
// list of files: a deck renders every slide with the same list.
const eastAsianFamilyCache = new Map<string, string | undefined>();

// The family resvg matches a face by: the typographic family when the face
// has one ("Hiragino Sans" for "Hiragino Sans W6"), else the family name.
const matchFamilyOf = (font: fontkit.Font): string =>
  font.getName('preferredFamily', 'en') ?? font.familyName;

// resvg skips a font file it cannot find or parse, and `fontFiles` promises
// the same, so such a file names no family here either.
const openFontFile = (path: string): fontkit.Font | fontkit.FontCollection | null => {
  if (!existsSync(path)) return null;
  try {
    return fontkit.create(readFileSync(path));
  } catch {
    return null;
  }
};

const eastAsianFamilyOf = (fontFiles: ReadonlyArray<string>): string | undefined => {
  const key = fontFiles.join('\n');
  if (eastAsianFamilyCache.has(key)) return eastAsianFamilyCache.get(key);
  let family: string | undefined;
  for (const path of fontFiles) {
    const opened = openFontFile(path);
    if (opened === null) continue;
    const faces = 'fonts' in opened ? opened.fonts : [opened];
    const face = faces.find((f) => f.hasGlyphForCodePoint(KANA_PROBE));
    if (face) {
      family = matchFamilyOf(face);
      break;
    }
  }
  eastAsianFamilyCache.set(key, family);
  return family;
};

const rasterize = (pres: PresentationData, slide: SlideData, opts: RenderImageOptions) => {
  // `getSlideSize` is only consulted to fail fast on a malformed deck; the SVG
  // carries its own viewBox and resvg fits to the requested width.
  getSlideSize(pres);
  const width = opts.width ?? DEFAULT_WIDTH;
  const measureText = opts.measureText ?? getSharedMeasurer();
  // resvg's own fallback for glyphs the run's family lacks ignores the weight
  // (bold Japanese paints with the regular face), so name the face instead.
  const eastAsianFamily = eastAsianFamilyOf(opts.fontFiles ?? []);
  const svg = renderSlideSvg(pres, slide, {
    measureText,
    textLayout: 'svg',
    ...(opts.resolveFamily === undefined ? {} : { resolveFamily: opts.resolveFamily }),
    ...(eastAsianFamily === undefined ? {} : { eastAsianFamily }),
  });
  const resvg = new Resvg(svg, {
    fitTo: { mode: 'width', value: width },
    font: {
      loadSystemFonts: opts.loadSystemFonts ?? false, // default: deterministic bundled set
      fontFiles: [...FONT_FILES, ...(opts.fontFiles ?? [])],
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
