// `getPresentationCommentCountsByAuthor(pres)` — deck-wide histogram
// of comment counts by author display name.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideComment,
  getPresentationCommentCountsByAuthor,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationCommentCountsByAuthor', () => {
  it('counts comments per author display name', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [slideA, slideB] = getSlides(pres);
    addSlideComment(slideA!, {
      author: { name: 'Alice', initials: 'A' },
      text: 'a1',
      position: { x: inches(0), y: inches(0) },
    });
    addSlideComment(slideA!, {
      author: { name: 'Alice', initials: 'A' },
      text: 'a2',
      position: { x: inches(0), y: inches(0) },
    });
    addSlideComment(slideB!, {
      author: { name: 'Bob', initials: 'B' },
      text: 'b1',
      position: { x: inches(0), y: inches(0) },
    });
    const counts = getPresentationCommentCountsByAuthor(pres);
    expect(counts.Alice).toBe(2);
    expect(counts.Bob).toBe(1);
  });

  it('returns an empty object on a deck with no comments', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getPresentationCommentCountsByAuthor(pres)).toEqual({});
  });
});
