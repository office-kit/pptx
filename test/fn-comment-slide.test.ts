// getCommentSlide — back-pointer from a comment to its owning slide.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideComment,
  findCommentsByAuthor,
  getCommentSlide,
  getSlideIndex,
  getSlides,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getCommentSlide', () => {
  it('round-trips back to the owning slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[1]!;
    const c = addSlideComment(slide, { author: { name: 'Alice' }, text: 'note' });
    expect(getCommentSlide(c)).toBe(slide);
    expect(getSlideIndex(pres, getCommentSlide(c))).toBe(getSlideIndex(pres, slide));
  });

  it('lets findCommentsByAuthor results report which slide they came from', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [first, second] = getSlides(pres);
    addSlideComment(first!, { author: { name: 'Alice' }, text: 'on 0' });
    addSlideComment(second!, { author: { name: 'Alice' }, text: 'on 1' });

    const hits = findCommentsByAuthor(pres, 'Alice');
    const slideIndices = hits.map((c) => getSlideIndex(pres, getCommentSlide(c))).sort();
    expect(slideIndices).toEqual([0, 1]);
  });
});
