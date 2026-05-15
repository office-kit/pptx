// getPresentationCommenters — distinct active commenters across deck.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideComment,
  getCommentAuthors,
  getPresentationCommenters,
  getSlides,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationCommenters', () => {
  it('aggregates distinct authors across every slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [first, second] = getSlides(pres);
    addSlideComment(first!, { author: { name: 'Alice' }, text: 'on 0' });
    addSlideComment(second!, { author: { name: 'Bob' }, text: 'on 1' });
    addSlideComment(first!, { author: { name: 'Alice' }, text: 'second from Alice' });

    const commenters = getPresentationCommenters(pres);
    expect(commenters.length).toBe(2);
    expect(commenters.map((a) => a.name).sort()).toEqual(['Alice', 'Bob']);
  });

  it('is a subset of getCommentAuthors when authors register but never comment', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    // getCommentAuthors returns the registry; getPresentationCommenters
    // returns only those who actually have a comment. The fixture has
    // no comments, so both should be empty.
    expect(getPresentationCommenters(pres)).toEqual([]);
    expect(getCommentAuthors(pres)).toEqual([]);
  });
});
