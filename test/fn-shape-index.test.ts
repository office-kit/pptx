// getShapeIndex — 0-based shape position on its slide.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getShapeIndex,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeIndex', () => {
  it('matches the position in getSlideShapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    for (const [i, shape] of getSlideShapes(slide).entries()) {
      expect(getShapeIndex(shape)).toBe(i);
    }
  });

  it('returns the new last-index after add', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const before = getSlideShapes(slide).length;
    const tb = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(2), h: inches(1), text: 'tail',
    });
    expect(getShapeIndex(tb)).toBe(before);
  });
});
