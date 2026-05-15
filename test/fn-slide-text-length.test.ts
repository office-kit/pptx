// getSlideTextLength — code-point length of all visible text.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  getSlideText,
  getSlideTextLength,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideTextLength', () => {
  it('matches Array.from(getSlideText).length', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(2), h: inches(1), text: 'hello',
    });
    addSlideTextBox(slide, {
      x: inches(0), y: inches(1), w: inches(2), h: inches(1), text: 'world',
    });
    expect(getSlideTextLength(slide)).toBe(Array.from(getSlideText(slide)).length);
  });

  it('counts emoji as one code point each', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout: blank });
    addSlideTextBox(slide, {
      // Two emoji — would be 4 in .length, but 2 as code points.
      x: inches(0), y: inches(0), w: inches(2), h: inches(1), text: '🎉🚀',
    });
    expect(getSlideTextLength(slide)).toBe(2);
  });
});
