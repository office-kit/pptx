// clearAllSlideComments — strip every comment across the deck.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideComment,
  clearAllSlideComments,
  findSlideLayout,
  getCommentCount,
  getSlides,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: clearAllSlideComments', () => {
  it('removes every comment and reports the count', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    addSlide(pres, { layout: blank });
    const slides = getSlides(pres);
    addSlideComment(slides[0]!, { author: { name: 'A' }, text: 'one' });
    addSlideComment(slides[1]!, { author: { name: 'A' }, text: 'two' });
    addSlideComment(slides[1]!, { author: { name: 'B' }, text: 'three' });
    expect(getCommentCount(pres)).toBe(3);
    expect(clearAllSlideComments(pres)).toBe(3);
    expect(getCommentCount(pres)).toBe(0);
  });

  it('returns 0 when no slide has comments', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(clearAllSlideComments(pres)).toBe(0);
  });
});
