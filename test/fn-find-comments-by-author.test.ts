// findCommentsByAuthor — filter every comment in the deck by author.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideComment,
  findCommentsByAuthor,
  getCommentAuthor,
  getCommentText,
  getSlides,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findCommentsByAuthor', () => {
  it('returns every comment authored by the matching name', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [first, second] = getSlides(pres);
    addSlideComment(first!, { author: { name: 'Alice' }, text: 'a1' });
    addSlideComment(first!, { author: { name: 'Bob' }, text: 'b1' });
    addSlideComment(second!, { author: { name: 'Alice' }, text: 'a2' });

    const alices = findCommentsByAuthor(pres, 'Alice');
    expect(alices.length).toBe(2);
    expect(alices.every((c) => getCommentAuthor(c).name === 'Alice')).toBe(true);
    expect(alices.map((c) => getCommentText(c)).sort()).toEqual(['a1', 'a2']);
  });

  it('returns an empty array for an unknown author', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    addSlideComment(getSlides(pres)[0]!, { author: { name: 'Alice' }, text: 'a' });
    expect(findCommentsByAuthor(pres, 'Carol')).toEqual([]);
  });

  it('is case-sensitive', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    addSlideComment(getSlides(pres)[0]!, { author: { name: 'Alice' }, text: 'a' });
    expect(findCommentsByAuthor(pres, 'alice')).toEqual([]);
  });
});
