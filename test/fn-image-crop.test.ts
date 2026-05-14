// Free-function image-crop API.
//
// Verifies `<a:srcRect>` is written / removed / merged correctly on the
// picture's `<p:blipFill>`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  Presentation,
  _internalPackageOf,
  getShapeKind,
  getSlideShapes,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeImageCrop,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await Presentation.load(bytes);
  const pkg = _internalPackageOf(pres);
  const part = pkg.parts.find((p) => p.name === `/ppt/slides/slide${slideIndex + 1}.xml`);
  if (!part) throw new Error(`slide${slideIndex + 1}.xml not found`);
  return new TextDecoder().decode(part.data);
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

  it('rejects out-of-range fractions', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => getShapeKind(s) === 'picture')!;
    expect(() => setShapeImageCrop(picture, { left: 1.5 })).toThrow();
    expect(() => setShapeImageCrop(picture, { top: -0.1 })).toThrow();
  });
});
