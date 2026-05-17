// findShapesAtPoint — every shape whose bounds contain (x, y).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  findShapesAtPoint,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const EMU_PER_INCH = 914400;

describe('fn API: findShapesAtPoint', () => {
  it('returns every shape whose bounds contain the point', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(3),
    });
    addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(3),
      h: inches(3),
    });
    // Point (2", 2") sits inside both rects.
    const hits = findShapesAtPoint(slide, 2 * EMU_PER_INCH, 2 * EMU_PER_INCH);
    expect(hits.length).toBe(2);
  });

  it('returns an empty list when no shape contains the point', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    expect(findShapesAtPoint(slide, 100 * EMU_PER_INCH, 100 * EMU_PER_INCH)).toEqual([]);
  });
});
