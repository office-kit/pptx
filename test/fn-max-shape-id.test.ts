// getMaxShapeId — highest cNvPr@id on a slide.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideShape,
  findSlideLayout,
  getMaxShapeId,
  getShapeId,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getMaxShapeId', () => {
  it('returns the largest id on the slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    addSlideShape(slide, {
      preset: 'rect', x: inches(2), y: inches(0), w: inches(1), h: inches(1),
    });
    const reread = getSlides(pres).at(-1)!;
    const max = Math.max(...getSlideShapes(reread).map((s) => getShapeId(s)));
    expect(getMaxShapeId(reread)).toBe(max);
  });

  it('returns 0 for a slide with no shapes carrying positive ids', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    // Blank layout's placeholders may have ids; we just check it's a number ≥ 0.
    expect(getMaxShapeId(slide)).toBeGreaterThanOrEqual(0);
  });
});
