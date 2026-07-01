// Picture contrast via the shared <a:blip><a:lum contrast="…"/>.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getShapeImageBrightness,
  getShapeImageContrast,
  getShapeKind,
  getSlideShapes,
  getSlides,
  loadPresentation,
  setShapeImageBrightness,
  setShapeImageContrast,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: setShapeImageContrast', () => {
  it('round-trips a contrast fraction', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => getShapeKind(s) === 'picture')!;
    expect(getShapeImageContrast(picture)).toBeNull();
    setShapeImageContrast(picture, 0.5);
    expect(getShapeImageContrast(picture)).toBeCloseTo(0.5);
    setShapeImageContrast(picture, -0.5);
    expect(getShapeImageContrast(picture)).toBeCloseTo(-0.5);
    setShapeImageContrast(picture, null);
    expect(getShapeImageContrast(picture)).toBeNull();
    setShapeImageContrast(picture, 0);
    expect(getShapeImageContrast(picture)).toBeNull(); // value 0 is identity → cleared
  });

  it('shares one <a:lum> element with brightness', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const picture = getSlideShapes(getSlides(pres)[0]!).find((s) => getShapeKind(s) === 'picture')!;
    setShapeImageBrightness(picture, 0.3);
    setShapeImageContrast(picture, -0.2);
    // Both corrections survive together.
    expect(getShapeImageBrightness(picture)).toBeCloseTo(0.3);
    expect(getShapeImageContrast(picture)).toBeCloseTo(-0.2);
    // Clearing one leaves the other intact.
    setShapeImageBrightness(picture, null);
    expect(getShapeImageBrightness(picture)).toBeNull();
    expect(getShapeImageContrast(picture)).toBeCloseTo(-0.2);
  });

  it('rejects non-pictures and out-of-range values', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const text = getSlideShapes(slide).find((s) => getShapeKind(s) === 'shape')!;
    expect(() => setShapeImageContrast(text, 0.5)).toThrow(/picture/);

    const pres2 = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const picture = getSlideShapes(getSlides(pres2)[0]!).find(
      (s) => getShapeKind(s) === 'picture',
    )!;
    expect(() => setShapeImageContrast(picture, -1.5)).toThrow(RangeError);
    expect(() => setShapeImageContrast(picture, 2.5)).toThrow(RangeError);
  });
});
