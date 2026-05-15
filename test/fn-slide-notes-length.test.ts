// getSlideNotesLength — code-point length of speaker notes.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  getSlideNotesLength,
  getSlides,
  loadPresentation,
  setSlideNotes,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideNotesLength', () => {
  it('counts code points, not UTF-16 units', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    const slide = getSlides(pres).at(-1)!;
    // Two emoji — would be 4 in .length, but 2 as code points.
    setSlideNotes(slide, '🎉🚀');
    expect(getSlideNotesLength(slide)).toBe(2);
  });

  it('returns 0 when no notes are present', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    for (const slide of getSlides(pres)) {
      expect(getSlideNotesLength(slide)).toBe(0);
    }
  });
});
