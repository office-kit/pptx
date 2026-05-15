// Paragraph spacing (spcBef / spcAft).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getParagraphSpacing,
  getSlides,
  inches,
  loadPresentation,
  setParagraphSpacing,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: setParagraphSpacing', () => {
  it('round-trips before + after spacing in points', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(3), h: inches(2),
      text: 'A\nB',
    });
    expect(getParagraphSpacing(tb, 1)).toEqual({ beforePts: null, afterPts: null });
    setParagraphSpacing(tb, 1, { beforePts: 6, afterPts: 3 });
    expect(getParagraphSpacing(tb, 1)).toEqual({ beforePts: 6, afterPts: 3 });
  });

  it('passing null on a side clears just that side', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(3), h: inches(2),
      text: 'A',
    });
    setParagraphSpacing(tb, 0, { beforePts: 4, afterPts: 2 });
    setParagraphSpacing(tb, 0, { beforePts: null });
    expect(getParagraphSpacing(tb, 0)).toEqual({ beforePts: null, afterPts: 2 });
  });

  it('rejects negative values', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(3), h: inches(2),
      text: 'A',
    });
    expect(() => setParagraphSpacing(tb, 0, { beforePts: -1 })).toThrow(RangeError);
  });
});
