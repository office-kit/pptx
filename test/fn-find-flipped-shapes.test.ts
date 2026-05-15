// findFlippedShapes — shapes flipped horizontally or vertically.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideShape,
  findFlippedShapes,
  findSlideLayout,
  getShapeFlip,
  getShapeId,
  inches,
  loadPresentation,
  setShapeFlip,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findFlippedShapes', () => {
  it('returns only flipped shapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    const flipped = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    setShapeFlip(flipped, { horizontal: true });
    addSlideShape(slide, {
      preset: 'rect', x: inches(2), y: inches(0), w: inches(1), h: inches(1),
    });
    const out = findFlippedShapes(slide);
    expect(out.length).toBe(1);
    expect(getShapeId(out[0]!)).toBe(getShapeId(flipped));
    expect(getShapeFlip(out[0]!)?.horizontal).toBe(true);
  });

  it('returns empty when no shape is flipped', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    expect(findFlippedShapes(slide)).toEqual([]);
  });
});
