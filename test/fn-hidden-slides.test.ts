// getHiddenSlides — every slide whose show="0" flag is set.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getHiddenSlides,
  getSlideIndex,
  getSlides,
  getVisibleSlides,
  loadPresentation,
  setSlideHidden,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getHiddenSlides', () => {
  it('returns an empty list when no slide is hidden', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getHiddenSlides(pres)).toEqual([]);
  });

  it('returns only slides marked hidden', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    setSlideHidden(slides[1]!, true);
    const hidden = getHiddenSlides(pres);
    expect(hidden.length).toBe(1);
    expect(getSlideIndex(pres, hidden[0]!)).toBe(1);
  });

  it('partitions every slide as either visible or hidden', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setSlideHidden(getSlides(pres)[0]!, true);
    const visible = getVisibleSlides(pres);
    const hidden = getHiddenSlides(pres);
    expect(visible.length + hidden.length).toBe(getSlides(pres).length);
  });
});
