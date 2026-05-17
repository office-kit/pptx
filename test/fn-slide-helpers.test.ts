// `clearSlideShapes` and `getSlideAt`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  clearSlideShapes,
  getSlideAt,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: clearSlideShapes', () => {
  it('removes every shape but keeps the slide structurally valid', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    addSlideShape(slide, {
      preset: 'ellipse',
      x: inches(0),
      y: inches(2),
      w: inches(1),
      h: inches(1),
    });
    const before = getSlideShapes(slide).length;
    expect(before).toBeGreaterThan(0);

    clearSlideShapes(slide);
    expect(getSlideShapes(slide).length).toBe(0);
  });
});

describe('fn API: getSlideAt', () => {
  it('returns the slide at a valid index', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    expect(getSlideAt(pres, 0)).toBe(slides[0]);
    expect(getSlideAt(pres, 1)).toBe(slides[1]);
  });

  it('returns null for out-of-range indices', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getSlideAt(pres, 99)).toBeNull();
    expect(getSlideAt(pres, -1)).toBeNull();
  });
});
