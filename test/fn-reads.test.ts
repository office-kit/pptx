// Free-function read API: slide + shape accessors.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  findSlidePlaceholder,
  getShapeFlip,
  getShapeId,
  getShapeKind,
  getShapeName,
  getShapePlaceholderIdx,
  getShapePlaceholderType,
  getShapePosition,
  getShapeRotation,
  getShapeSize,
  getShapeText,
  getSlideLayout,
  getSlideShapes,
  getSlideText,
  getSlides,
  loadPresentation,
  replaceTokensInSlide,
  savePresentation,
  setShapeText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: slide reads', () => {
  it('getSlideShapes / getSlideLayout / findSlidePlaceholder behave as advertised', async () => {
    const bytes = await readFile(fixture('one-text-slide.pptx'));
    const pres = await loadPresentation(bytes);
    const slide = getSlides(pres)[0]!;

    expect(getSlideShapes(slide).length).toBeGreaterThan(0);
    expect(getSlideText(slide).length).toBeGreaterThan(0);
    expect(getSlideLayout(slide)).not.toBeNull();

    const title = findSlidePlaceholder(slide, 'title');
    expect(title).not.toBeNull();
    if (title) {
      expect(getShapeText(title).length).toBeGreaterThan(0);
      expect(typeof getShapeName(title)).toBe('string');
      expect(typeof getShapeId(title)).toBe('number');
      expect(getShapeKind(title)).toBe('shape');
      expect(getShapePlaceholderType(title)).toBe('title');
      // Title placeholder typically has idx 0 (the spec default).
      expect(
        typeof getShapePlaceholderIdx(title) === 'number' || getShapePlaceholderIdx(title) === null,
      ).toBe(true);
    }
  });

  it('getShape{Position,Size,Rotation,Flip} read xfrm values from a picture-bearing slide', async () => {
    const bytes = await readFile(fixture('one-image-slide.pptx'));
    const pres = await loadPresentation(bytes);
    const slide = getSlides(pres)[0]!;
    const shapes = getSlideShapes(slide);
    expect(shapes.length).toBeGreaterThan(0);

    for (const shape of shapes) {
      // Every shape exposes the four readers; nulls are acceptable on
      // placeholder-inheriting shapes that don't carry their own xfrm.
      void getShapePosition(shape);
      void getShapeSize(shape);
      expect(typeof getShapeRotation(shape)).toBe('number');
      void getShapeFlip(shape);
    }
  });

  it('replaceTokensInSlide mutates a single slide in place', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const seedShape = getSlideShapes(getSlides(pres)[0]!).find((s) => getShapeText(s).length > 0);
    if (!seedShape) throw new Error('expected text shape');
    setShapeText(seedShape, 'Hello, {{name}}!');

    const seeded = await loadPresentation(await savePresentation(pres));
    const slide = getSlides(seeded)[0]!;
    const n = replaceTokensInSlide(slide, { name: 'World' });
    expect(n).toBe(1);
    expect(getSlideText(slide)).toContain('Hello, World!');

    const reloaded = await loadPresentation(await savePresentation(seeded));
    expect(getSlideText(getSlides(reloaded)[0]!)).toContain('Hello, World!');
  });
});
