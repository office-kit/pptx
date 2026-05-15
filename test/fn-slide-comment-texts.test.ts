// getSlideCommentTexts — per-slide list of comment text.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideComment,
  findSlideLayout,
  getCommentText,
  getSlideCommentTexts,
  getSlideComments,
  getSlides,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideCommentTexts', () => {
  it('matches getSlideComments(...).map(getCommentText)', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideComment(slide, { author: { name: 'A' }, text: 'first' });
    addSlideComment(slide, { author: { name: 'B' }, text: 'second' });
    const reread = getSlides(pres).at(-1)!;
    expect(getSlideCommentTexts(reread)).toEqual(
      getSlideComments(reread).map((c) => getCommentText(c)),
    );
    expect(getSlideCommentTexts(reread)).toContain('first');
    expect(getSlideCommentTexts(reread)).toContain('second');
  });

  it('returns empty for a slide without comments', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    expect(getSlideCommentTexts(slide)).toEqual([]);
  });
});
