// `getPresentationCommentCountsBySlide(pres)` — dense per-slide
// comment count histogram.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideComment,
  getPresentationCommentCountsBySlide,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationCommentCountsBySlide', () => {
  it('returns 0 for every slide on a clean deck', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getPresentationCommentCountsBySlide(pres)).toEqual([0, 0]);
  });

  it('counts comments per slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [slideA, slideB] = getSlides(pres);
    addSlideComment(slideA!, {
      author: { name: 'A', initials: 'a' },
      text: 'a1',
      position: { x: inches(0), y: inches(0) },
    });
    addSlideComment(slideA!, {
      author: { name: 'A', initials: 'a' },
      text: 'a2',
      position: { x: inches(0), y: inches(0) },
    });
    addSlideComment(slideB!, {
      author: { name: 'B', initials: 'b' },
      text: 'b1',
      position: { x: inches(0), y: inches(0) },
    });
    expect(getPresentationCommentCountsBySlide(pres)).toEqual([2, 1]);
  });

  it('array length matches the slide count', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getPresentationCommentCountsBySlide(pres).length).toBe(getSlides(pres).length);
  });
});
