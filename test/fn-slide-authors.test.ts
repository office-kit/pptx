// getSlideAuthors — distinct authors who have commented on a slide.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideComment,
  getSlideAuthors,
  getSlides,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideAuthors', () => {
  it('lists each distinct author once', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideComment(slide, { author: { name: 'Alice' }, text: 'a1' });
    addSlideComment(slide, { author: { name: 'Bob' }, text: 'b1' });
    addSlideComment(slide, { author: { name: 'Alice' }, text: 'a2' });

    const authors = getSlideAuthors(slide);
    expect(authors.length).toBe(2);
    expect(authors.map((a) => a.name).sort()).toEqual(['Alice', 'Bob']);
  });

  it('returns an empty list when no comments exist', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getSlideAuthors(getSlides(pres)[0]!)).toEqual([]);
  });

  it('preserves first-seen author order', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideComment(slide, { author: { name: 'Charlie' }, text: 'c' });
    addSlideComment(slide, { author: { name: 'Alice' }, text: 'a' });
    addSlideComment(slide, { author: { name: 'Charlie' }, text: 'c2' });

    const authors = getSlideAuthors(slide);
    expect(authors.map((a) => a.name)).toEqual(['Charlie', 'Alice']);
  });
});
