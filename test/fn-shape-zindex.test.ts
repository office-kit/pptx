// getShapeZIndex + setShapeZIndex.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getShapeName,
  getShapeZIndex,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  setShapeZIndex,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const orderedNames = (slide: ReturnType<typeof getSlides>[number]): string[] =>
  getSlideShapes(slide).map((s) => getShapeName(s));

describe('fn API: shape z-index', () => {
  it('getShapeZIndex returns the document-order position', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const a = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'A',
    });
    const b = addSlideShape(slide, {
      preset: 'ellipse',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'B',
    });
    expect(getShapeZIndex(a)).toBeGreaterThan(-1);
    expect(getShapeZIndex(b)).toBeGreaterThan(getShapeZIndex(a));
  });

  it('setShapeZIndex moves the shape to the requested position', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'A',
    });
    addSlideShape(slide, {
      preset: 'ellipse',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'B',
    });
    const c = addSlideShape(slide, {
      preset: 'triangle',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'C',
    });

    setShapeZIndex(c, 0);
    expect(orderedNames(slide)[0]).toBe('C');
  });

  it('setShapeZIndex clamps to the available range', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const a = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'A',
    });
    setShapeZIndex(a, 999);
    // After clamp the shape sits at the last position among shapes.
    expect(orderedNames(slide).at(-1)).toBe('A');
  });
});
