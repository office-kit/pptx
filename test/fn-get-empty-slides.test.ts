// `getEmptySlides(pres)` — slides with no shapes.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  getEmptySlides,
  getSlideIndex,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getEmptySlides', () => {
  it('does not return slides that carry shapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const populated = addBlankSlide(pres);
    addSlideTextBox(populated, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'hi',
    });
    const found = getEmptySlides(pres);
    expect(found.map((s) => getSlideIndex(pres, s))).not.toContain(getSlideIndex(pres, populated));
  });

  it('result length is at most the deck length', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getEmptySlides(pres).length).toBeLessThanOrEqual(getSlides(pres).length);
  });
});
