// `RenderImageOptions.resolveFamily` lets a PNG render draw a deck's font in a
// family the caller loaded, instead of the bundled substitute every authored
// name otherwise maps to, so the raster matches a measurer that knows that font.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  inches,
  loadPresentation,
  setShapeTextFormat,
} from '@office-kit/pptx';
import {
  defaultMeasurer,
  renderSlideToSvg,
  substituteFamily,
  type FontSpec,
  type TextMeasurer,
} from '../packages/preview/src/index.ts';
import { renderSlideToRgba, type RgbaImage } from '../packages/preview/src/node.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const AUTHORED = 'Meiryo';
// A bundled face standing in for one the caller adds through `fontFiles`, so
// the test needs no font outside the repository.
const LOADED = 'Liberation Mono';
const EAST_ASIAN = 'Test Sans JP';
const WIDTH = 960;
// Channel value below which a pixel counts as text ink on the white slide.
const INK_THRESHOLD = 128;

const resolveAuthored = (family: string | null): string =>
  family === AUTHORED ? LOADED : substituteFamily(family);

const deck = async (text: string) => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(8),
    h: inches(1),
    text,
  });
  setShapeTextFormat(shape, { font: AUTHORED, size: 28 });
  return { pres, slide };
};

const inkBox = (image: RgbaImage) => {
  let ink = 0;
  let minX = image.width;
  let maxX = -1;
  for (let i = 0; i < image.data.length; i += 4) {
    if (image.data[i]! < INK_THRESHOLD) {
      ink++;
      const x = (i / 4) % image.width;
      minX = Math.min(minX, x);
      maxX = Math.max(maxX, x);
    }
  }
  return { ink, width: maxX - minX + 1 };
};

const recordingMeasurer = (): { measure: TextMeasurer; families: Set<string> } => {
  const families = new Set<string>();
  return {
    families,
    measure: (text: string, spec: FontSpec) => {
      families.add(spec.family);
      return defaultMeasurer(text, spec);
    },
  };
};

describe('renderSlideToRgba resolveFamily', () => {
  it('draws the run in the resolved family instead of the bundled substitute', async () => {
    const { pres, slide } = await deck('Quarterly revenue grew');

    const substituted = inkBox(renderSlideToRgba(pres, slide, { width: WIDTH }).image);
    const resolved = inkBox(
      renderSlideToRgba(pres, slide, { width: WIDTH, resolveFamily: resolveAuthored }).image,
    );

    // Liberation Mono's 0.6 em advance is far wider than Carlito's letters.
    expect(resolved.width).toBeGreaterThan(substituted.width * 1.2);
    expect(resolved.ink).not.toBe(substituted.ink);
  });

  it('measures the run with the resolved family', async () => {
    const { pres, slide } = await deck('Quarterly revenue grew');
    const { measure, families } = recordingMeasurer();

    renderSlideToRgba(pres, slide, {
      width: WIDTH,
      measureText: measure,
      resolveFamily: resolveAuthored,
    });

    expect(families).toContain(LOADED);
    expect(families).not.toContain(substituteFamily(AUTHORED));
  });

  it('leaves the render unchanged when unset', async () => {
    const { pres, slide } = await deck('Quarterly revenue grew');

    const unset = renderSlideToRgba(pres, slide, { width: WIDTH }).png;
    const substitute = renderSlideToRgba(pres, slide, {
      width: WIDTH,
      resolveFamily: substituteFamily,
    }).png;

    expect(Buffer.from(substitute).equals(Buffer.from(unset))).toBe(true);
  });
});

describe('resolveFamily with eastAsianFamily', () => {
  it('writes East Asian characters in eastAsianFamily and the rest in the resolved family', async () => {
    const { pres, slide } = await deck('売上 Q3 は（前年比）+12%');
    const { measure, families } = recordingMeasurer();

    const svg = renderSlideToSvg(pres, slide, {
      textLayout: 'svg',
      measureText: measure,
      resolveFamily: resolveAuthored,
      eastAsianFamily: EAST_ASIAN,
    });
    const spans = [...svg.matchAll(/<tspan([^>]*)>([^<]*)<\/tspan>/g)].map((m) => [
      m[2],
      /font-family="([^"]*)"/.exec(m[1]!)?.[1],
    ]);

    expect(spans).toEqual([
      ['売上', EAST_ASIAN],
      [' Q3 ', LOADED],
      ['は（前年比）', EAST_ASIAN],
      ['+12%', LOADED],
    ]);
    expect([...families].sort()).toEqual([EAST_ASIAN, LOADED].sort());
  });
});
