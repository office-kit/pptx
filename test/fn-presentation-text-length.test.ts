// getPresentationTextLength — total visible code-point text across the deck.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  getPresentationTextLength,
  getSlideTextLength,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationTextLength', () => {
  it('equals the sum of per-slide text lengths', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'alpha',
    });
    addSlide(pres, { layout: blank });
    const expected = getSlides(pres)
      .map((s) => getSlideTextLength(s))
      .reduce((a, b) => a + b, 0);
    expect(getPresentationTextLength(pres)).toBe(expected);
  });

  it('counts surrogate-pair emoji as one code point', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: '🎉🚀',
    });
    expect(getPresentationTextLength(pres)).toBe(2);
  });
});
