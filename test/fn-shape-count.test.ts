// getShapeCount / getTotalShapeCount — slide-level + deck-level
// shape counters.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getShapeCount,
  getSlideShapes,
  getSlides,
  getTotalShapeCount,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeCount / getTotalShapeCount', () => {
  it('matches getSlideShapes().length per slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    for (const slide of getSlides(pres)) {
      expect(getShapeCount(slide)).toBe(getSlideShapes(slide).length);
    }
  });

  it('total equals sum across slides', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const expected = getSlides(pres).reduce((sum, s) => sum + getShapeCount(s), 0);
    expect(getTotalShapeCount(pres)).toBe(expected);
  });

  it('tracks additions', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const before = getTotalShapeCount(pres);
    addSlideShape(getSlides(pres)[0]!, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    addSlideShape(getSlides(pres)[1]!, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    expect(getTotalShapeCount(pres)).toBe(before + 2);
  });
});
