// getSlideTextShapeCount — fast counter of text-bearing shapes.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideShape,
  addSlideTextBox,
  findShapesWithText,
  findSlideLayout,
  getSlideTextShapeCount,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideTextShapeCount', () => {
  it('matches findShapesWithText length', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1), text: 'a',
    });
    addSlideTextBox(slide, {
      x: inches(1), y: inches(0), w: inches(1), h: inches(1), text: 'b',
    });
    addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(1), w: inches(1), h: inches(1),
    });
    expect(getSlideTextShapeCount(slide)).toBe(findShapesWithText(slide).length);
    expect(getSlideTextShapeCount(slide)).toBeGreaterThanOrEqual(2);
  });

  it('returns 0 on a fresh blank slide with no added text', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    // Blank layouts may still carry empty dt/sldNum/ftr placeholders;
    // those have empty text and don't count as text-bearing.
    expect(getSlideTextShapeCount(slide)).toBe(0);
    // Also matches the original fixture slides
    const first = getSlides(pres)[0]!;
    expect(getSlideTextShapeCount(first)).toBe(findShapesWithText(first).length);
  });
});
