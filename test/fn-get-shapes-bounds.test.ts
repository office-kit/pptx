// getShapesBounds — union bounding box across multiple shapes.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getShapesBounds,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapesBounds', () => {
  it('computes the union of two non-overlapping shapes', async () => {
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
      x: inches(3),
      y: inches(1),
      w: inches(2),
      h: inches(2),
    });

    const bounds = getShapesBounds([a, b])!;
    expect(bounds.x).toBe(0);
    expect(bounds.y).toBe(0);
    expect(bounds.w).toBe(5 * 914400); // 0 → 5in
    expect(bounds.h).toBe(3 * 914400); // 0 → 3in
  });

  it('returns null for an empty list', async () => {
    expect(getShapesBounds([])).toBeNull();
  });
});
