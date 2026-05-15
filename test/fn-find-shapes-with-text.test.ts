// findShapesWithText — every shape on the slide carrying text.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  addSlideTextBox,
  findShapesWithText,
  getSlides,
  hasShapeText,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findShapesWithText', () => {
  it('returns only shapes with non-empty text', async () => {
    const { addSlide, findSlideLayout } = await import('../src/api/index.ts');
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(2), h: inches(1), text: 'one',
    });
    addSlideTextBox(slide, {
      x: inches(0), y: inches(1), w: inches(2), h: inches(1), text: 'two',
    });
    addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(2), w: inches(1), h: inches(1),
    });

    const shapes = findShapesWithText(slide);
    expect(shapes.length).toBe(2);
    for (const s of shapes) expect(hasShapeText(s)).toBe(true);
    void getSlides;
  });
});
