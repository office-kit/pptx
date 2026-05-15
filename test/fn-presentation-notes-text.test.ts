// getPresentationNotesText — joined speaker notes across the deck.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  getPresentationNotesText,
  getSlides,
  loadPresentation,
  setSlideNotes,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationNotesText', () => {
  it('joins notes with the default form-feed separator', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    addSlide(pres, { layout: blank });
    const slides = getSlides(pres);
    setSlideNotes(slides[0]!, 'alpha');
    setSlideNotes(slides[1]!, 'beta');
    expect(getPresentationNotesText(pres)).toContain('alpha\fbeta');
  });

  it('uses the provided separator', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    addSlide(pres, { layout: blank });
    const slides = getSlides(pres);
    setSlideNotes(slides[0]!, 'A');
    setSlideNotes(slides[1]!, 'B');
    expect(getPresentationNotesText(pres, '||')).toContain('A||B');
  });

  it('is empty when no slide has notes', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    // Default join is just separators between empty strings, so no letters.
    const joined = getPresentationNotesText(pres);
    expect(joined.replace(/\f/g, '')).toBe('');
  });
});
