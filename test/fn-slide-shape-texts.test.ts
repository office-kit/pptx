// getSlideShapeTexts — every shape's visible text in document order.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  addSlideTextBox,
  getShapeText,
  getSlideShapeTexts,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideShapeTexts', () => {
  it('matches getSlideShapes(...).map(getShapeText)', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1), text: 'alpha',
    });
    addSlideTextBox(slide, {
      x: inches(0), y: inches(1), w: inches(1), h: inches(1), text: 'beta',
    });
    const expected = getSlideShapes(slide).map((s) => getShapeText(s));
    expect(getSlideShapeTexts(slide)).toEqual(expected);
  });

  it('returns empty string for shapes with no text body', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    const texts = getSlideShapeTexts(slide);
    // The shape we just added carries no text, so at least one entry is ''.
    expect(texts).toContain('');
  });
});
