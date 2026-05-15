// getSlideShapeIds — every shape's numeric id in document order.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getShapeId,
  getSlideShapeIds,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideShapeIds', () => {
  it('matches getSlideShapes(...).map(getShapeId)', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1), text: 'a',
    });
    addSlideTextBox(slide, {
      x: inches(0), y: inches(1), w: inches(1), h: inches(1), text: 'b',
    });
    const expected = getSlideShapes(slide).map((s) => getShapeId(s));
    expect(getSlideShapeIds(slide)).toEqual(expected);
  });

  it('returns unique ids', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1), text: 'a',
    });
    const ids = getSlideShapeIds(slide);
    expect(new Set(ids).size).toBe(ids.length);
  });
});
