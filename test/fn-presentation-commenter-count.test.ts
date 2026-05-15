// getPresentationCommenterCount — fast count of distinct reviewers.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideComment,
  findSlideLayout,
  getPresentationCommenterCount,
  getPresentationCommenters,
  getSlides,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationCommenterCount', () => {
  it('matches getPresentationCommenters length', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    addSlide(pres, { layout: blank });
    const slides = getSlides(pres);
    addSlideComment(slides[0]!, { author: { name: 'Alice' }, text: '1' });
    addSlideComment(slides[1]!, { author: { name: 'Bob' }, text: '2' });
    addSlideComment(slides[1]!, { author: { name: 'Alice' }, text: '3' });
    expect(getPresentationCommenterCount(pres)).toBe(
      getPresentationCommenters(pres).length,
    );
    expect(getPresentationCommenterCount(pres)).toBe(2);
  });

  it('returns 0 when no comments exist', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getPresentationCommenterCount(pres)).toBe(0);
  });
});
