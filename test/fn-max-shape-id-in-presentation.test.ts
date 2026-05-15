// getMaxShapeIdInPresentation — deck-wide highest cNvPr@id.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideShape,
  findSlideLayout,
  getMaxShapeId,
  getMaxShapeIdInPresentation,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getMaxShapeIdInPresentation', () => {
  it('matches the max of getMaxShapeId across slides', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    addSlide(pres, { layout: blank });
    {
      const slides = getSlides(pres);
      addSlideShape(slides[0]!, {
        preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
      });
      addSlideShape(slides[1]!, {
        preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
      });
    }
    const slides = getSlides(pres);
    const perSlideMax = Math.max(...slides.map((s) => getMaxShapeId(s)));
    expect(getMaxShapeIdInPresentation(pres)).toBe(perSlideMax);
  });

  it('returns a non-negative number on a fresh deck', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getMaxShapeIdInPresentation(pres)).toBeGreaterThanOrEqual(0);
  });
});
