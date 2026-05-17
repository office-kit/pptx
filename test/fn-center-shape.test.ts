// centerShapeOnSlide — move a shape to the slide canvas center.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  centerShapeOnSlide,
  getShapeBounds,
  getSlideSize,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: centerShapeOnSlide', () => {
  it('moves the shape so its center matches the slide center', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const rect = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(2),
    });
    const size = getSlideSize(pres)!;
    centerShapeOnSlide(rect);

    const bounds = getShapeBounds(rect)!;
    // Center at (slideW/2 - shapeW/2, slideH/2 - shapeH/2).
    expect(bounds.x).toBe(Math.round(size.width / 2 - bounds.w / 2));
    expect(bounds.y).toBe(Math.round(size.height / 2 - bounds.h / 2));
  });
});
