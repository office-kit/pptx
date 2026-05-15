// findSlideByTitle — exact-match lookup against the title placeholder.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideByTitle,
  findSlideLayout,
  getSlideTitle,
  getSlides,
  loadPresentation,
  setSlideTitle,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findSlideByTitle', () => {
  it('matches the title exactly', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    // Add slides up-front; their handles go stale after subsequent
    // addSlide invalidates the slides cache, so we compare by title.
    setSlideTitle(addSlide(pres, { layout }), 'Quarterly Results');
    setSlideTitle(addSlide(pres, { layout }), 'Roadmap');

    const found = findSlideByTitle(pres, 'Quarterly Results');
    expect(found).not.toBeNull();
    expect(getSlideTitle(found!)).toBe('Quarterly Results');
  });

  it('returns null on a near-miss', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    setSlideTitle(addSlide(pres, { layout }), 'Quarterly Results');
    // Substring would match; exact equality does not.
    expect(findSlideByTitle(pres, 'Quarterly')).toBeNull();
    expect(findSlideByTitle(pres, 'Roadmap')).toBeNull();
  });

  it('matches the first slide when there are duplicates', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    setSlideTitle(addSlide(pres, { layout }), 'Duplicate');
    setSlideTitle(addSlide(pres, { layout }), 'Duplicate');

    const found = findSlideByTitle(pres, 'Duplicate');
    const slides = getSlides(pres);
    // Should be the first slide bearing that title — index 0 (or
    // wherever the first "Duplicate" lands, but always before any
    // later one with the same title).
    expect(found).toBe(slides.find((s) => getSlideTitle(s) === 'Duplicate'));
  });
});
