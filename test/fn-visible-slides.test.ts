// getVisibleSlides — slides whose <p:sldId show="0"> flag is not set.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getSlideIndex,
  getSlides,
  getVisibleSlides,
  loadPresentation,
  setSlideHidden,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getVisibleSlides', () => {
  it('returns every slide when none are hidden', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const visible = getVisibleSlides(pres);
    expect(visible.length).toBe(getSlides(pres).length);
  });

  it('skips slides marked hidden', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    setSlideHidden(slides[0]!, true);
    const visible = getVisibleSlides(pres);
    expect(visible.length).toBe(slides.length - 1);
    expect(visible.every((s) => getSlideIndex(pres, s) !== 0)).toBe(true);
  });

  it('restores visibility when the flag is cleared', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    setSlideHidden(slides[0]!, true);
    setSlideHidden(slides[0]!, false);
    expect(getVisibleSlides(pres).length).toBe(slides.length);
  });
});
