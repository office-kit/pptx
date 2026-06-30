// Picture brightness via the shared <a:blip><a:lum bright="…"/>.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getShapeImageBrightness,
  getShapeKind,
  getSlideShapes,
  getSlides,
  loadPresentation,
  setShapeImageBrightness,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: setShapeImageBrightness', () => {
  it('round-trips a brightness fraction', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => getShapeKind(s) === 'picture')!;
    expect(getShapeImageBrightness(picture)).toBeNull();
    setShapeImageBrightness(picture, 0.4);
    expect(getShapeImageBrightness(picture)).toBeCloseTo(0.4);
    setShapeImageBrightness(picture, -0.2);
    expect(getShapeImageBrightness(picture)).toBeCloseTo(-0.2);
    setShapeImageBrightness(picture, null);
    expect(getShapeImageBrightness(picture)).toBeNull();
  });

  it('rejects non-pictures and out-of-range values', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const text = getSlideShapes(slide).find((s) => getShapeKind(s) === 'shape')!;
    expect(() => setShapeImageBrightness(text, 0.5)).toThrow(/picture/);

    const pres2 = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const picture = getSlideShapes(getSlides(pres2)[0]!).find(
      (s) => getShapeKind(s) === 'picture',
    )!;
    expect(() => setShapeImageBrightness(picture, 1.5)).toThrow(RangeError);
    expect(() => setShapeImageBrightness(picture, -1.5)).toThrow(RangeError);
  });
});
