// getSlideShapeKinds — every shape's kind in document order.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  addSlideTextBox,
  getShapeKind,
  getSlideShapeKinds,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideShapeKinds', () => {
  it('matches getSlideShapes(...).map(getShapeKind)', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    addSlideTextBox(slide, {
      x: inches(0), y: inches(1), w: inches(2), h: inches(1), text: 'hi',
    });
    const expected = getSlideShapes(slide).map((s) => getShapeKind(s));
    expect(getSlideShapeKinds(slide)).toEqual(expected);
  });
});
