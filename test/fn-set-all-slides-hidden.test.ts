// setAllSlidesHidden — bulk hide/show every slide.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getSlides,
  isSlideHidden,
  loadPresentation,
  setAllSlidesHidden,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: setAllSlidesHidden', () => {
  it('hides every slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setAllSlidesHidden(pres, true);
    for (const slide of getSlides(pres)) {
      expect(isSlideHidden(slide)).toBe(true);
    }
  });

  it('reveals every slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setAllSlidesHidden(pres, true);
    setAllSlidesHidden(pres, false);
    for (const slide of getSlides(pres)) {
      expect(isSlideHidden(slide)).toBe(false);
    }
  });
});
