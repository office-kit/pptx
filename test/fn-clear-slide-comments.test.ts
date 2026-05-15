// clearSlideComments — slide-scoped comment stripper.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideComment,
  clearSlideComments,
  findSlideLayout,
  getSlideCommentCount,
  getSlides,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: clearSlideComments', () => {
  it('clears only the target slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    addSlide(pres, { layout: blank });
    let slides = getSlides(pres);
    addSlideComment(slides[0]!, { author: { name: 'A' }, text: '1' });
    addSlideComment(slides[1]!, { author: { name: 'A' }, text: '2' });
    addSlideComment(slides[1]!, { author: { name: 'B' }, text: '3' });
    slides = getSlides(pres);
    expect(clearSlideComments(slides[0]!)).toBe(1);
    slides = getSlides(pres);
    expect(getSlideCommentCount(slides[0]!)).toBe(0);
    expect(getSlideCommentCount(slides[1]!)).toBe(2);
  });

  it('returns 0 when slide has nothing to clear', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    expect(clearSlideComments(slide)).toBe(0);
  });
});
