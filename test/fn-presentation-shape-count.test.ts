// getPresentationShapeCount — total shape count across the deck.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  addSlideTextBox,
  getAllShapes,
  getPresentationShapeCount,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationShapeCount', () => {
  it('matches getAllShapes length', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [first, second] = getSlides(pres);
    addSlideShape(first!, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    addSlideTextBox(second!, {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1), text: 'b',
    });
    expect(getPresentationShapeCount(pres)).toBe(getAllShapes(pres).length);
  });

  it('grows when a shape is added', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const before = getPresentationShapeCount(pres);
    addSlideShape(getSlides(pres)[0]!, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    expect(getPresentationShapeCount(pres)).toBe(before + 1);
  });
});
