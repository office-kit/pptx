// Picture contrast via <a:lumMod>.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getShapeImageContrast,
  getShapeKind,
  getSlideShapes,
  getSlides,
  loadPresentation,
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
    setShapeImageContrast(picture, 1.5);
    expect(getShapeImageContrast(picture)).toBeCloseTo(1.5);
    setShapeImageContrast(picture, null);
    expect(getShapeImageContrast(picture)).toBeNull();
    setShapeImageContrast(picture, 1);
    expect(getShapeImageContrast(picture)).toBeNull(); // value 1 is identity → cleared
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
    expect(() => setShapeImageContrast(picture, -0.1)).toThrow(RangeError);
    expect(() => setShapeImageContrast(picture, 2.5)).toThrow(RangeError);
  });
});
