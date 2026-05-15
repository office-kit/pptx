// getSlideShapeBounds — every shape's bounds in document order.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  addSlideTextBox,
  getShapeBounds,
  getSlideShapeBounds,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideShapeBounds', () => {
  it('matches getSlideShapes(...).map(getShapeBounds)', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    addSlideTextBox(slide, {
      x: inches(2), y: inches(2), w: inches(3), h: inches(1), text: 't',
    });
    const expected = getSlideShapes(slide).map((s) => getShapeBounds(s));
    expect(getSlideShapeBounds(slide)).toEqual(expected);
  });

  it('returns matching length with shape array', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(getSlideShapeBounds(slide).length).toBe(getSlideShapes(slide).length);
  });
});
