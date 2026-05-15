// getHiddenSlideCount / getVisibleSlideCount — fast counters.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getHiddenSlideCount,
  getHiddenSlides,
  getSlides,
  getVisibleSlideCount,
  getVisibleSlides,
  loadPresentation,
  setSlideHidden,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getHiddenSlideCount / getVisibleSlideCount', () => {
  it('matches array-length forms', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getHiddenSlideCount(pres)).toBe(getHiddenSlides(pres).length);
    expect(getVisibleSlideCount(pres)).toBe(getVisibleSlides(pres).length);
  });

  it('updates after flipping a slide hidden', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const before = getHiddenSlideCount(pres);
    setSlideHidden(getSlides(pres)[0]!, true);
    expect(getHiddenSlideCount(pres)).toBe(before + 1);
    expect(getHiddenSlideCount(pres) + getVisibleSlideCount(pres))
      .toBe(getSlides(pres).length);
  });
});
