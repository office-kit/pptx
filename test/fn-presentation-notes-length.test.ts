// getPresentationNotesLength — total code-point length of notes.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  getPresentationNotesLength,
  getSlideNotesLength,
  getSlides,
  loadPresentation,
  setSlideNotes,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationNotesLength', () => {
  it('equals the sum of per-slide notes lengths', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    addSlide(pres, { layout: blank });
    const slides = getSlides(pres);
    setSlideNotes(slides[0]!, 'alpha');
    setSlideNotes(slides[1]!, 'beta');
    const expected = slides.map((s) => getSlideNotesLength(s)).reduce((a, b) => a + b, 0);
    expect(getPresentationNotesLength(pres)).toBe(expected);
  });

  it('counts surrogate-pair emoji as one code point', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    const slide = getSlides(pres).at(-1)!;
    setSlideNotes(slide, '🎉🚀');
    expect(getPresentationNotesLength(pres)).toBe(2);
  });
});
