// getSlidesWithComments — slides carrying at least one comment.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideComment,
  getSlideIndex,
  getSlidesWithComments,
  getSlides,
  loadPresentation,
  removeSlideComment,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlidesWithComments', () => {
  it('returns an empty list on a comment-free deck', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getSlidesWithComments(pres)).toEqual([]);
  });

  it('returns only the slides that have comments', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [first, second] = getSlides(pres);
    addSlideComment(first!, { author: { name: 'A' }, text: 'note' });
    void second;
    const hits = getSlidesWithComments(pres);
    expect(hits.length).toBe(1);
    expect(getSlideIndex(pres, hits[0]!)).toBe(0);
  });

  it('drops a slide once its last comment is removed', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const c = addSlideComment(slide, { author: { name: 'A' }, text: 'note' });
    expect(getSlidesWithComments(pres).length).toBe(1);
    removeSlideComment(c);
    expect(getSlidesWithComments(pres)).toEqual([]);
  });
});
