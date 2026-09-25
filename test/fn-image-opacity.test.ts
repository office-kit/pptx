// Picture opacity via `<a:alphaModFix>`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import {
  getShapeKind,
  getShapeImageBytes,
  getShapeImageOpacity,
  setShapeImageFill,
  getSlideShapes,
  getSlideXmlString,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeImageOpacity,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await loadPresentation(bytes);
  return getSlideXmlString(getSlides(pres)[slideIndex]!);
};

describe('fn API: setShapeImageOpacity', () => {
  it('writes <a:alphaModFix amt="..."/> with the converted ST_Percentage', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => getShapeKind(s) === 'picture')!;
    setShapeImageOpacity(picture, 0.5);
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('<a:alphaModFix amt="50000"/>');
  });

  it('passing null clears any existing alphaModFix', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => getShapeKind(s) === 'picture')!;
    setShapeImageOpacity(picture, 0.25);
    expect(await slideXml(await savePresentation(pres), 0)).toContain('alphaModFix');
    setShapeImageOpacity(picture, null);
    expect(await slideXml(await savePresentation(pres), 0)).not.toContain('alphaModFix');
  });

  it('rejects non-picture shapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const text = getSlideShapes(slide).find((s) => getShapeKind(s) === 'shape')!;
    expect(() => setShapeImageOpacity(text, 0.5)).toThrow(/picture/);
  });

  it('rejects out-of-range opacities', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => getShapeKind(s) === 'picture')!;
    expect(() => setShapeImageOpacity(picture, 1.5)).toThrow(RangeError);
    expect(() => setShapeImageOpacity(picture, -0.1)).toThrow(RangeError);
  });
  it('preserves existing opacity when rejecting an invalid value', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => getShapeKind(s) === 'picture')!;
    setShapeImageOpacity(picture, 0.25);
    const before = getSlideXmlString(slide);
    for (const value of [-1, 2, NaN, Infinity]) {
      expect(() => setShapeImageOpacity(picture, value)).toThrow(RangeError);
      expect(getShapeImageOpacity(picture)).toBe(0.25);
      expect(getSlideXmlString(slide)).toBe(before);
    }
  });

  it('round-trips opacity on a regular shape with an image fill', async () => {
    const images = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const picture = getSlideShapes(getSlides(images)[0]!).find(
      (s) => getShapeKind(s) === 'picture',
    )!;
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide).find((s) => getShapeKind(s) === 'shape')!;
    setShapeImageFill(shape, getShapeImageBytes(picture)!);
    setShapeImageOpacity(shape, 0.35);
    expect(getShapeImageOpacity(shape)).toBe(0.35);
    expect(renderSlideToSvg(pres, slide)).toMatch(/<image[^>]*opacity="0.350"/);
    const restored = await loadPresentation(await savePresentation(pres));
    const restoredShape = getSlideShapes(getSlides(restored)[0]!).find(
      (s) => getShapeKind(s) === 'shape',
    )!;
    expect(getShapeImageOpacity(restoredShape)).toBe(0.35);
    setShapeImageOpacity(restoredShape, null);
    expect(getShapeImageOpacity(restoredShape)).toBeNull();
  });
});
