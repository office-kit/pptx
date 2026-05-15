// getShapeAnimation — read back the v1 single-effect animation state.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  clearSlideAnimations,
  getShapeAnimation,
  getSlideShapes,
  getSlides,
  loadPresentation,
  setShapeAnimation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeAnimation', () => {
  it('returns null when no animation is set', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    expect(getShapeAnimation(shape)).toBeNull();
  });

  it('returns the configured preset for every supported effect', async () => {
    for (const effect of ['appear', 'fadeIn', 'disappear', 'fadeOut'] as const) {
      const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
      const slide = getSlides(pres)[0]!;
      const shape = getSlideShapes(slide)[0]!;
      setShapeAnimation(shape, { effect });
      expect(getShapeAnimation(shape)).toBe(effect);
    }
  });

  it('returns null after clearSlideAnimations', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeAnimation(shape, { effect: 'fadeIn' });
    clearSlideAnimations(slide);
    expect(getShapeAnimation(shape)).toBeNull();
  });
});
