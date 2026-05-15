// getSlideCommentAuthors — distinct authors on a slide.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideComment,
  findSlideLayout,
  getSlideCommentAuthors,
  getSlides,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideCommentAuthors', () => {
  it('dedupes by author id and preserves first-seen order', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideComment(slide, { author: { name: 'Alice' }, text: '1' });
    addSlideComment(slide, { author: { name: 'Bob' }, text: '2' });
    addSlideComment(slide, { author: { name: 'Alice' }, text: '3' });
    const reread = getSlides(pres).at(-1)!;
    const names = getSlideCommentAuthors(reread).map((a) => a.name);
    expect(names).toEqual(['Alice', 'Bob']);
  });

  it('returns empty for a slide without comments', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    expect(getSlideCommentAuthors(slide)).toEqual([]);
  });
});
