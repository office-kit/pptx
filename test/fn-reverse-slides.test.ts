// reverseSlides — flip the slide order.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  getSlideTitle,
  getSlides,
  loadPresentation,
  reverseSlides,
  setSlideTitle,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: reverseSlides', () => {
  it('flips a three-slide deck', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    setSlideTitle(addSlide(pres, { layout }), 'A');
    setSlideTitle(addSlide(pres, { layout }), 'B');
    setSlideTitle(addSlide(pres, { layout }), 'C');
    reverseSlides(pres);
    expect(getSlides(pres).map((s) => getSlideTitle(s))).toEqual(['C', 'B', 'A']);
  });

  it('is a no-op on a 0-slide deck', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    expect(() => reverseSlides(pres)).not.toThrow();
    expect(getSlides(pres)).toEqual([]);
  });

  it('is its own inverse', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    setSlideTitle(addSlide(pres, { layout }), 'A');
    setSlideTitle(addSlide(pres, { layout }), 'B');
    setSlideTitle(addSlide(pres, { layout }), 'C');
    reverseSlides(pres);
    reverseSlides(pres);
    expect(getSlides(pres).map((s) => getSlideTitle(s))).toEqual(['A', 'B', 'C']);
  });
});
