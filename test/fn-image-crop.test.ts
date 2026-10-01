// Free-function image-crop API.
//
// Verifies `<a:srcRect>` is written / removed / merged correctly on the
// picture's `<p:blipFill>`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import {
  getShapeKind,
  getShapeImageCrop,
  setShapeRotation,
  getSlideShapes,
  getSlideXmlString,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeImageCrop,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await loadPresentation(bytes);
  return getSlideXmlString(getSlides(pres)[slideIndex]!);
};

describe('fn API: setShapeImageCrop', () => {
  it('writes a srcRect with the converted ST_Percentage values', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => getShapeKind(s) === 'picture');
    if (!picture) throw new Error('expected picture');

    setShapeImageCrop(picture, { left: 0.1, top: 0.2, right: 0.15, bottom: 0.05 });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('<a:srcRect');
    expect(xml).toContain('l="10000"');
    expect(xml).toContain('t="20000"');
    expect(xml).toContain('r="15000"');
    expect(xml).toContain('b="5000"');
  });

  it('round-trips source outsets and signed percentage limits', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const picture = getSlideShapes(getSlides(pres)[0]!).find((s) => getShapeKind(s) === 'picture')!;
    for (const crop of [
      { left: -0.25, top: -0.5, right: 0.1, bottom: 0 },
      { left: -21474.83648, top: 0, right: 21474.83647, bottom: 1.5 },
    ]) {
      setShapeImageCrop(picture, crop);
      const reloaded = await loadPresentation(await savePresentation(pres));
      const restored = getSlideShapes(getSlides(reloaded)[0]!).find(
        (s) => getShapeKind(s) === 'picture',
      )!;
      expect(getShapeImageCrop(restored)).toEqual(crop);
    }
  });

  it('omits sides that are zero', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => getShapeKind(s) === 'picture')!;
    setShapeImageCrop(picture, { left: 0.25 });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('l="25000"');
    expect(xml).not.toContain('t="0"');
    expect(xml).not.toMatch(/<a:srcRect[^>]+r="/);
  });

  it('passing null removes any existing srcRect', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => getShapeKind(s) === 'picture')!;
    setShapeImageCrop(picture, { left: 0.3, right: 0.3 });
    expect(await slideXml(await savePresentation(pres), 0)).toContain('<a:srcRect');
    setShapeImageCrop(picture, null);
    expect(await slideXml(await savePresentation(pres), 0)).not.toContain('<a:srcRect');
  });

  it('throws for non-picture shapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const textShape = getSlideShapes(slide).find((s) => getShapeKind(s) === 'shape')!;
    expect(() => setShapeImageCrop(textShape, { left: 0.1 })).toThrow(/picture/);
  });

  it('preserves the existing crop after invalid input and a subsequent edit', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => getShapeKind(s) === 'picture')!;
    setShapeImageCrop(picture, { left: 0.2, bottom: 0.1 });
    const before = getShapeImageCrop(picture);
    expect(() => setShapeImageCrop(picture, { left: 0.3, bottom: Number.NaN })).toThrow();
    expect(getShapeImageCrop(picture)).toEqual(before);
    setShapeRotation(picture, 10);
    const reloaded = await loadPresentation(await savePresentation(pres));
    const restored = getSlideShapes(getSlides(reloaded)[0]!).find(
      (s) => getShapeKind(s) === 'picture',
    )!;
    expect(getShapeImageCrop(restored)).toEqual(before);
  });

  it('rejects out-of-range fractions', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => getShapeKind(s) === 'picture')!;
    expect(() => setShapeImageCrop(picture, { left: 21474.83648 })).toThrow();
    expect(() => setShapeImageCrop(picture, { top: -21474.83649 })).toThrow();
  });
});

it('reads percentage-suffixed image crops without changing their magnitude', async () => {
  const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
  const picture = getSlideShapes(getSlides(pres)[0]!).find((s) => getShapeKind(s) === 'picture')!;
  setShapeImageCrop(picture, { left: -0.125, top: 0.25, right: 0.375, bottom: -0.5 });
  const files = unzipSync(await savePresentation(pres));
  const path = 'ppt/slides/slide1.xml';
  files[path] = strToU8(
    strFromU8(files[path]!).replace(
      /<a:srcRect[^>]*\/>/,
      '<a:srcRect l="-12.5%" t="25%" r="37.5%" b="-50%"/>',
    ),
  );
  const loaded = await loadPresentation(zipSync(files));
  const restored = getSlideShapes(getSlides(loaded)[0]!).find(
    (s) => getShapeKind(s) === 'picture',
  )!;
  expect(getShapeImageCrop(restored)).toEqual({
    left: -0.125,
    top: 0.25,
    right: 0.375,
    bottom: -0.5,
  });
  const reloaded = await loadPresentation(await savePresentation(loaded));
  expect(
    getShapeImageCrop(
      getSlideShapes(getSlides(reloaded)[0]!).find((s) => getShapeKind(s) === 'picture')!,
    ),
  ).toEqual(getShapeImageCrop(restored));
});
