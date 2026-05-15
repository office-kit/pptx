// clearAllAnimations — wipe animations from every slide.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  clearAllAnimations,
  getShapeAnimation,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  setShapeAnimation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: clearAllAnimations', () => {
  it('clears every slide\'s animation timing', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [first, second] = getSlides(pres);
    const s1 = addSlideShape(first!, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    const s2 = addSlideShape(second!, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    setShapeAnimation(s1, { effect: 'fadeIn' });
    setShapeAnimation(s2, { effect: 'fadeIn' });

    clearAllAnimations(pres);

    // Re-fetch shapes after cache refresh from clearSlideAnimations.
    const slides = getSlides(pres);
    for (const slide of slides) {
      for (const shape of getSlideShapes(slide)) {
        expect(getShapeAnimation(shape)).toBeNull();
      }
    }
  });

  it('is a no-op on a deck with no animations', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(() => clearAllAnimations(pres)).not.toThrow();
  });
});
