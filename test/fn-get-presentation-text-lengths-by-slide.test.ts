// `getPresentationTextLengthsBySlide(pres)` — dense per-slide text
// length histogram.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getPresentationTextLengthsBySlide,
  getSlides,
  getSlideTextLength,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationTextLengthsBySlide', () => {
  it('matches getSlideTextLength for every slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const arr = getPresentationTextLengthsBySlide(pres);
    const slides = getSlides(pres);
    expect(arr.length).toBe(slides.length);
    for (let i = 0; i < slides.length; i++) {
      expect(arr[i]).toBe(getSlideTextLength(slides[i]!));
    }
  });

  it('responds to new text after mutation', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const before = getPresentationTextLengthsBySlide(pres);
    const slideA = getSlides(pres)[0]!;
    addSlideTextBox(slideA, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'hello',
    });
    const after = getPresentationTextLengthsBySlide(pres);
    // Adding the textbox raises the slide's text length by at least the
    // string itself; PowerPoint inserts an inter-paragraph separator
    // between text bodies so we don't pin to an exact delta.
    expect(after[0]!).toBeGreaterThanOrEqual(before[0]! + 'hello'.length);
  });
});
