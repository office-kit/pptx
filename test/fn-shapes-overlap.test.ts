// shapesOverlap — axis-aligned bounding-box overlap test.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getSlides,
  inches,
  loadPresentation,
  shapesOverlap,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: shapesOverlap', () => {
  it('detects overlap', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const a = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(2),
    });
    const b = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
    });
    expect(shapesOverlap(a, b)).toBe(true);
  });

  it('reports no overlap for disjoint shapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const a = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    const b = addSlideShape(slide, {
      preset: 'rect',
      x: inches(5),
      y: inches(5),
      w: inches(1),
      h: inches(1),
    });
    expect(shapesOverlap(a, b)).toBe(false);
  });

  it('reports no overlap for edge-touching shapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const a = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    // b sits exactly to the right of a, sharing one edge — not an
    // overlap with the strict-inequality definition.
    const b = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    expect(shapesOverlap(a, b)).toBe(false);
  });
});
