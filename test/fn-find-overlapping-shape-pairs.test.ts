// findOverlappingShapePairs — unordered pairs of bounds-overlapping shapes.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideShape,
  findOverlappingShapePairs,
  findSlideLayout,
  getShapeId,
  inches,
  loadPresentation,
  shapesOverlap,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findOverlappingShapePairs', () => {
  it('returns each colliding pair exactly once', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    const a = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(3), h: inches(3),
    });
    const b = addSlideShape(slide, {
      preset: 'rect', x: inches(1), y: inches(1), w: inches(3), h: inches(3),
    });
    // Disjoint
    addSlideShape(slide, {
      preset: 'rect', x: inches(10), y: inches(10), w: inches(1), h: inches(1),
    });

    const pairs = findOverlappingShapePairs(slide);
    expect(pairs.length).toBe(1);
    expect(shapesOverlap(pairs[0]![0], pairs[0]![1])).toBe(true);
    // Document-order: a's id strictly precedes b's id within the pair.
    expect(getShapeId(pairs[0]![0])).toBe(getShapeId(a));
    expect(getShapeId(pairs[0]![1])).toBe(getShapeId(b));
  });

  it('returns empty when no shapes overlap', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    addSlideShape(slide, {
      preset: 'rect', x: inches(3), y: inches(3), w: inches(1), h: inches(1),
    });
    expect(findOverlappingShapePairs(slide)).toEqual([]);
  });
});
