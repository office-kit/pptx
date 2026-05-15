// getShapeAt — indexed shape access.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getShapeAt,
  getShapeIndex,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeAt', () => {
  it('matches positional access via getSlideShapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    for (const slide of getSlides(pres)) {
      const shapes = getSlideShapes(slide);
      for (let i = 0; i < shapes.length; i++) {
        expect(getShapeAt(slide, i)).toBe(shapes[i]);
      }
    }
  });

  it('returns null for negative and out-of-range indices', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(getShapeAt(slide, -1)).toBeNull();
    expect(getShapeAt(slide, 9999)).toBeNull();
  });

  it('round-trips against getShapeIndex', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    for (const shape of getSlideShapes(slide)) {
      const idx = getShapeIndex(shape);
      expect(getShapeAt(slide, idx)).toBe(shape);
    }
  });
});
