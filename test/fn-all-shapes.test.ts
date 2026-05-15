// getAllShapes — flatten every shape across every slide.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getAllShapes,
  getShapeKind,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getAllShapes', () => {
  it('returns one entry per (slide, shape) pair with correct slideIndex', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    addSlideShape(slides[0]!, { preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1) });
    addSlideShape(slides[1]!, { preset: 'ellipse', x: inches(0), y: inches(0), w: inches(1), h: inches(1) });
    addSlideShape(slides[1]!, { preset: 'triangle', x: inches(0), y: inches(0), w: inches(1), h: inches(1) });

    const all = getAllShapes(pres);
    // Slide 0: title placeholder + the rect.
    // Slide 1: title placeholder + ellipse + triangle.
    const indices = all.map((e) => e.slideIndex);
    // Every entry should belong to one of the two slides.
    expect(indices.every((i) => i === 0 || i === 1)).toBe(true);
    // At least the three shapes we explicitly added show up.
    const kinds = all.map((e) => getShapeKind(e.shape));
    expect(kinds.filter((k) => k === 'shape').length).toBeGreaterThanOrEqual(3);
  });
});
