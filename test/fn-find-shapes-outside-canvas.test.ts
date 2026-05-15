// findShapesOutsideCanvas — shapes that overflow the slide canvas.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideShape,
  findShapesOutsideCanvas,
  findSlideLayout,
  getShapeBounds,
  getSlideSize,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findShapesOutsideCanvas', () => {
  it('flags shapes whose bounds extend past the canvas', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    const size = getSlideSize(pres)!;
    // In-canvas shape
    addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    // Way off the right edge
    addSlideShape(slide, {
      preset: 'rect', x: size.width, y: inches(0), w: inches(1), h: inches(1),
    });
    const offscreen = findShapesOutsideCanvas(slide, pres);
    expect(offscreen.length).toBe(1);
    const b = getShapeBounds(offscreen[0]!)!;
    expect(b.x).toBeGreaterThanOrEqual(size.width);
  });

  it('returns empty when every shape is within the canvas', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    expect(findShapesOutsideCanvas(slide, pres)).toEqual([]);
  });
});
