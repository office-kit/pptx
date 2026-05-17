// pointInShape — hit test against a shape's axis-aligned bounds.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getSlides,
  inches,
  loadPresentation,
  pointInShape,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const EMU_PER_INCH = 914400;

describe('fn API: pointInShape', () => {
  it('reports true for points inside the shape', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const rect = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
    });
    // Center of the rect.
    expect(pointInShape(rect, 2 * EMU_PER_INCH, 2 * EMU_PER_INCH)).toBe(true);
  });

  it('reports false for points outside', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const rect = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
    });
    expect(pointInShape(rect, 0, 0)).toBe(false);
    expect(pointInShape(rect, 10 * EMU_PER_INCH, 10 * EMU_PER_INCH)).toBe(false);
  });

  it('uses half-open semantics on the bottom-right edge', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const rect = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
    });
    // Top-left edge: inclusive.
    expect(pointInShape(rect, EMU_PER_INCH, EMU_PER_INCH)).toBe(true);
    // Bottom-right corner: exclusive.
    expect(pointInShape(rect, 3 * EMU_PER_INCH, 3 * EMU_PER_INCH)).toBe(false);
  });
});
