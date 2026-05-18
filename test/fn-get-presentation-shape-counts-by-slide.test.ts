// `getPresentationShapeCountsBySlide(pres)` — dense per-slide shape
// count histogram.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getPresentationShapeCountsBySlide,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationShapeCountsBySlide', () => {
  it("reports each slide's shape count in order", async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [slideA, slideB] = getSlides(pres);
    const baseline = getPresentationShapeCountsBySlide(pres);
    addSlideTextBox(slideA!, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'A1',
    });
    addSlideTextBox(slideA!, {
      x: inches(0),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'A2',
    });
    addSlideTextBox(slideB!, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'B1',
    });
    const counts = getPresentationShapeCountsBySlide(pres);
    expect(counts[0]).toBe(baseline[0]! + 2);
    expect(counts[1]).toBe(baseline[1]! + 1);
  });

  it('array length matches the slide count', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getPresentationShapeCountsBySlide(pres).length).toBe(getSlides(pres).length);
  });
});
