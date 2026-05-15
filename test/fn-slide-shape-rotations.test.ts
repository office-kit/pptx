// getSlideShapeRotations — each shape's rotation in degrees.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideShape,
  findSlideLayout,
  getShapeRotation,
  getSlideShapeRotations,
  getSlideShapes,
  inches,
  loadPresentation,
  setShapeRotation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideShapeRotations', () => {
  it('matches getSlideShapes(...).map(getShapeRotation)', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    const a = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    setShapeRotation(a, 45);
    addSlideShape(slide, {
      preset: 'rect', x: inches(2), y: inches(0), w: inches(1), h: inches(1),
    });
    const expected = getSlideShapes(slide).map((s) => getShapeRotation(s));
    expect(getSlideShapeRotations(slide)).toEqual(expected);
  });

  it('reports 0 for un-rotated shapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    // Every shape on a freshly-added blank slide is un-rotated.
    for (const r of getSlideShapeRotations(slide)) expect(r).toBe(0);
  });
});
