// getSlideCommentCount — count comments on a single slide.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideComment,
  getSlideCommentCount,
  getSlides,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideCommentCount', () => {
  it('matches the size of getSlideComments', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(getSlideCommentCount(slide)).toBe(0);
    addSlideComment(slide, { author: { name: 'A' }, text: 'first' });
    addSlideComment(slide, { author: { name: 'A' }, text: 'second' });
    expect(getSlideCommentCount(slide)).toBe(2);
  });
});
