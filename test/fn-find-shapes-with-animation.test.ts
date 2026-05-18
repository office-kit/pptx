// `findShapesWithAnimation(slide)` — slide-scoped audit for shapes
// that have an authored animation effect.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  findShapesWithAnimation,
  getSlideShapes,
  getSlides,
  loadPresentation,
  setShapeAnimation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findShapesWithAnimation', () => {
  it('returns shapes that have an animation effect', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeAnimation(shape, { effect: 'fadeIn' });
    const matches = findShapesWithAnimation(slide);
    expect(matches.length).toBe(1);
  });

  it('returns an empty array on a slide with no animations', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(findShapesWithAnimation(slide)).toEqual([]);
  });
});
