// findSlidesWithOverlap — deck-wide audit for colliding shapes.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideShape,
  findSlideLayout,
  findSlidesWithOverlap,
  getSlideIndex,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findSlidesWithOverlap', () => {
  it('returns only slides with colliding shapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    // Slide A: overlapping pair
    {
      const overlap = addSlide(pres, { layout: blank });
      addSlideShape(overlap, {
        preset: 'rect', x: inches(0), y: inches(0), w: inches(3), h: inches(3),
      });
      addSlideShape(overlap, {
        preset: 'rect', x: inches(1), y: inches(1), w: inches(3), h: inches(3),
      });
    }
    const overlapIdx = 0; // first added by us
    // Slide B: disjoint pair
    {
      const clean = addSlide(pres, { layout: blank });
      addSlideShape(clean, {
        preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
      });
      addSlideShape(clean, {
        preset: 'rect', x: inches(3), y: inches(3), w: inches(1), h: inches(1),
      });
    }
    const cleanIdx = 1; // second added by us

    const flagged = findSlidesWithOverlap(pres);
    const indices = flagged.map((s) => getSlideIndex(pres, s));
    expect(indices).toContain(overlapIdx);
    expect(indices).not.toContain(cleanIdx);
  });

  it('returns empty when no slide has overlap', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    expect(findSlidesWithOverlap(pres)).toEqual([]);
  });
});
