// getShapeSlide — back-pointer from a shape to its owning slide.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  findShapeInPresentation,
  getShapeSlide,
  getSlideIndex,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeSlide', () => {
  it('returns the owning slide for a newly-added shape', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[1]!;
    const s = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'hi',
    });
    expect(getShapeSlide(s)).toBe(slide);
    expect(getSlideIndex(pres, getShapeSlide(s))).toBe(getSlideIndex(pres, slide));
  });

  it('lets a presentation-wide search resolve back to its slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[1]!;
    addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'unique-marker',
      name: 'MarkerShape',
    });

    const hit = findShapeInPresentation(pres, 'MarkerShape');
    expect(hit).not.toBeNull();
    expect(getShapeSlide(hit!)).toBe(slide);
  });
});
