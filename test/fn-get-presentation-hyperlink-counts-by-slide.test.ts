// `getPresentationHyperlinkCountsBySlide(pres)` — dense per-slide
// hyperlink count array.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getPresentationHyperlinkCountsBySlide,
  getSlides,
  inches,
  loadPresentation,
  setShapeHyperlink,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationHyperlinkCountsBySlide', () => {
  it('counts hyperlinks per slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [slideA, slideB] = getSlides(pres);
    const tA1 = addSlideTextBox(slideA!, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'A1',
    });
    const tA2 = addSlideTextBox(slideA!, {
      x: inches(0),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'A2',
    });
    const tB1 = addSlideTextBox(slideB!, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'B1',
    });
    setShapeHyperlink(tA1, 'https://a1.example/');
    setShapeHyperlink(tA2, 'https://a2.example/');
    setShapeHyperlink(tB1, 'https://b1.example/');
    expect(getPresentationHyperlinkCountsBySlide(pres)).toEqual([2, 1]);
  });

  it('returns all zeros when no shape has a hyperlink', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getPresentationHyperlinkCountsBySlide(pres)).toEqual([0, 0]);
  });
});
