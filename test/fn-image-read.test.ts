// Read-back for picture opacity + crop.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getShapeImageCrop,
  getShapeImageOpacity,
  getShapeKind,
  getSlideShapes,
  getSlides,
  loadPresentation,
  setShapeImageCrop,
  setShapeImageOpacity,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: image opacity read-back', () => {
  it('returns null when no override is set', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => getShapeKind(s) === 'picture')!;
    expect(getShapeImageOpacity(picture)).toBeNull();
  });

  it('round-trips an opacity fraction', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => getShapeKind(s) === 'picture')!;
    setShapeImageOpacity(picture, 0.5);
    expect(getShapeImageOpacity(picture)).toBeCloseTo(0.5);
    setShapeImageOpacity(picture, null);
    expect(getShapeImageOpacity(picture)).toBeNull();
  });
});

describe('fn API: image crop read-back', () => {
  it('returns null when no crop is set', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => getShapeKind(s) === 'picture')!;
    expect(getShapeImageCrop(picture)).toBeNull();
  });

  it('round-trips per-side crop fractions', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => getShapeKind(s) === 'picture')!;
    setShapeImageCrop(picture, { left: 0.1, top: 0.2, right: 0.05, bottom: 0.15 });
    const c = getShapeImageCrop(picture);
    expect(c).not.toBeNull();
    expect(c!.left).toBeCloseTo(0.1);
    expect(c!.top).toBeCloseTo(0.2);
    expect(c!.right).toBeCloseTo(0.05);
    expect(c!.bottom).toBeCloseTo(0.15);
  });
});
