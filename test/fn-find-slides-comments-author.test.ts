// findSlidesWithCommentsByAuthor — slides with reviewer-specific notes.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideComment,
  findSlidesWithCommentsByAuthor,
  getSlideIndex,
  getSlides,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findSlidesWithCommentsByAuthor', () => {
  it('lists slides where the author has any comment', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [first, second] = getSlides(pres);
    addSlideComment(first!, { author: { name: 'Alice' }, text: 'on 0' });
    addSlideComment(second!, { author: { name: 'Bob' }, text: 'on 1' });
    addSlideComment(second!, { author: { name: 'Alice' }, text: 'also on 1' });

    const alicesSlides = findSlidesWithCommentsByAuthor(pres, 'Alice');
    expect(alicesSlides.length).toBe(2);
    expect(alicesSlides.map((s) => getSlideIndex(pres, s)).sort()).toEqual([0, 1]);

    const bobsSlides = findSlidesWithCommentsByAuthor(pres, 'Bob');
    expect(bobsSlides.length).toBe(1);
  });

  it('returns an empty list for an unknown author', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    addSlideComment(getSlides(pres)[0]!, { author: { name: 'A' }, text: 'a' });
    expect(findSlidesWithCommentsByAuthor(pres, 'no-one')).toEqual([]);
  });
});
