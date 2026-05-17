// slideHasAnimations — predicate for "does this slide carry <p:timing>?"

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getSlideShapes,
  getSlides,
  loadPresentation,
  setShapeAnimation,
  slideHasAnimations,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: slideHasAnimations', () => {
  it('returns false for slides without <p:timing>', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    for (const slide of getSlides(pres)) {
      expect(slideHasAnimations(slide)).toBe(false);
    }
  });

  it('flips true after setShapeAnimation', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0];
    if (!shape) return;
    expect(slideHasAnimations(slide)).toBe(false);
    setShapeAnimation(shape, { effect: 'fadeIn' });
    expect(slideHasAnimations(slide)).toBe(true);
  });
});
