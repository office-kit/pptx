// clearAllSlideBackgrounds — wipe slide-level background overrides.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  clearAllSlideBackgrounds,
  getSlideBackground,
  getSlides,
  loadPresentation,
  setSlideBackground,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: clearAllSlideBackgrounds', () => {
  it('removes background overrides across the deck', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    for (const slide of getSlides(pres)) {
      setSlideBackground(slide, '#FF8800');
    }
    for (const slide of getSlides(pres)) {
      expect(getSlideBackground(slide).kind).toBe('solid');
    }

    clearAllSlideBackgrounds(pres);

    for (const slide of getSlides(pres)) {
      const bg = getSlideBackground(slide);
      expect(bg.kind).toBe('inherit');
    }
  });

  it('is a no-op on a deck with no overrides', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(() => clearAllSlideBackgrounds(pres)).not.toThrow();
  });
});
