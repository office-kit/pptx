// findRotatedShapes — shapes with non-zero rotation.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideShape,
  findRotatedShapes,
  findSlideLayout,
  getShapeId,
  getShapeRotation,
  inches,
  loadPresentation,
  setShapeRotation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findRotatedShapes', () => {
  it('returns only shapes with non-zero rotation', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    const rotated = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    setShapeRotation(rotated, 30);
    addSlideShape(slide, {
      preset: 'rect', x: inches(2), y: inches(0), w: inches(1), h: inches(1),
    });
    const out = findRotatedShapes(slide);
    expect(out.length).toBe(1);
    expect(getShapeId(out[0]!)).toBe(getShapeId(rotated));
    expect(getShapeRotation(out[0]!)).toBe(30);
  });

  it('returns empty when no shape is rotated', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    expect(findRotatedShapes(slide)).toEqual([]);
  });
});
