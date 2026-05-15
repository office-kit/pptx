// getSlideCount + getSlideLayoutCount — cheap O(1)-ish counters.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  getSlideCount,
  getSlideLayoutCount,
  getSlideLayouts,
  getSlides,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideCount + getSlideLayoutCount', () => {
  it('matches getSlides().length and getSlideLayouts().length', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getSlideCount(pres)).toBe(getSlides(pres).length);
    expect(getSlideLayoutCount(pres)).toBe(getSlideLayouts(pres).length);
  });

  it('tracks slide additions', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const before = getSlideCount(pres);
    addSlide(pres, { layout });
    addSlide(pres, { layout });
    expect(getSlideCount(pres)).toBe(before + 2);
  });

  it('returns the same count regardless of whether the cache is warm', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    // Cold cache.
    const cold = getSlideCount(pres);
    // Warm cache via getSlides.
    void getSlides(pres);
    const warm = getSlideCount(pres);
    expect(cold).toBe(warm);
  });
});
