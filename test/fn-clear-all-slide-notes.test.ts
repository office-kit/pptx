// clearAllSlideNotes — strip every slide's speaker notes.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  clearAllSlideNotes,
  findSlideLayout,
  getNotesSlideCount,
  getSlides,
  hasSlideNotes,
  loadPresentation,
  setSlideNotes,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: clearAllSlideNotes', () => {
  it('drops every notes part and reports the count', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    addSlide(pres, { layout: blank });
    const slides = getSlides(pres);
    setSlideNotes(slides[0]!, 'internal comment');
    setSlideNotes(slides[1]!, 'launch on monday');
    expect(getNotesSlideCount(pres)).toBe(2);
    expect(clearAllSlideNotes(pres)).toBe(2);
    expect(getNotesSlideCount(pres)).toBe(0);
    for (const slide of getSlides(pres)) {
      expect(hasSlideNotes(slide)).toBe(false);
    }
  });

  it('returns 0 when no slide has notes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(clearAllSlideNotes(pres)).toBe(0);
  });
});
