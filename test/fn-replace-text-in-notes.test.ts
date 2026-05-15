// replaceTextInNotes — bulk substring/regex replace across notes.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  getSlideNotes,
  getSlides,
  loadPresentation,
  replaceTextInNotes,
  setSlideNotes,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: replaceTextInNotes', () => {
  it('substring replace updates every matching slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    addSlide(pres, { layout: blank });
    let slides = getSlides(pres);
    setSlideNotes(slides[0]!, 'Acme launches at noon');
    setSlideNotes(slides[1]!, 'Q&A for Acme demo');
    expect(replaceTextInNotes(pres, 'Acme', 'Globex')).toBe(2);
    slides = getSlides(pres);
    expect(getSlideNotes(slides[0]!)).toBe('Globex launches at noon');
    expect(getSlideNotes(slides[1]!)).toBe('Q&A for Globex demo');
  });

  it('RegExp replace replaces all occurrences', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    const slides = getSlides(pres);
    setSlideNotes(slides[0]!, 'date 2025-12-01 and date 2026-01-15');
    expect(replaceTextInNotes(pres, /\d{4}-\d{2}-\d{2}/, 'YYYY-MM-DD')).toBe(1);
    expect(getSlideNotes(getSlides(pres)[0]!)).toBe('date YYYY-MM-DD and date YYYY-MM-DD');
  });

  it('returns 0 when nothing matches', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    setSlideNotes(getSlides(pres)[0]!, 'plain text');
    expect(replaceTextInNotes(pres, 'no-such-token', 'x')).toBe(0);
  });
});
