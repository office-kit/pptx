// getNotesSlideCount — fast counter of slides with notes.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  getNotesSlideCount,
  getSlides,
  getSlidesWithNotes,
  loadPresentation,
  setSlideNotes,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getNotesSlideCount', () => {
  it('matches getSlidesWithNotes length', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    addSlide(pres, { layout: blank });
    addSlide(pres, { layout: blank });
    const slides = getSlides(pres);
    setSlideNotes(slides[0]!, 'one');
    setSlideNotes(slides[2]!, 'three');
    expect(getNotesSlideCount(pres)).toBe(getSlidesWithNotes(pres).length);
    expect(getNotesSlideCount(pres)).toBeGreaterThanOrEqual(2);
  });

  it('returns 0 when no slide has notes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getNotesSlideCount(pres)).toBe(0);
  });
});
