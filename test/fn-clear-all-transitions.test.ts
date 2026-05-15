// clearAllTransitions — wipe transitions across every slide.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  clearAllTransitions,
  getSlideTransition,
  getSlides,
  loadPresentation,
  setSlideTransition,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: clearAllTransitions', () => {
  it('clears every slide\'s transition', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    for (const slide of getSlides(pres)) {
      setSlideTransition(slide, { effect: 'fade' });
    }
    for (const slide of getSlides(pres)) {
      expect(getSlideTransition(slide)).not.toBeNull();
    }

    clearAllTransitions(pres);

    for (const slide of getSlides(pres)) {
      expect(getSlideTransition(slide)).toBeNull();
    }
  });

  it('is a no-op on a deck with no transitions', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(() => clearAllTransitions(pres)).not.toThrow();
  });
});
