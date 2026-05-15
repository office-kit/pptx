// getAllComments — every comment paired with its slide index.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideComment,
  getAllComments,
  getCommentText,
  getSlides,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getAllComments', () => {
  it('returns an empty list on a deck without comments', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getAllComments(pres)).toEqual([]);
  });

  it('pairs each comment with its slide index', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [first, second] = getSlides(pres);
    addSlideComment(first!, { author: { name: 'A' }, text: '0-a' });
    addSlideComment(second!, { author: { name: 'A' }, text: '1-a' });
    addSlideComment(second!, { author: { name: 'B' }, text: '1-b' });

    const entries = getAllComments(pres);
    expect(entries.length).toBe(3);
    expect(entries[0]!.slideIndex).toBe(0);
    expect(getCommentText(entries[0]!.comment)).toBe('0-a');
    expect(entries[1]!.slideIndex).toBe(1);
    expect(entries[2]!.slideIndex).toBe(1);
  });
});
