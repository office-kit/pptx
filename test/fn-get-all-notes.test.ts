// getAllNotes — every slide's speaker notes paired with its index.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  getAllNotes,
  getSlides,
  loadPresentation,
  setSlideNotes,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getAllNotes', () => {
  it('returns an empty list on a deck without notes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getAllNotes(pres)).toEqual([]);
  });

  it('pairs each note with its slide index', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout });
    addSlide(pres, { layout });
    addSlide(pres, { layout });
    const slides = getSlides(pres);
    setSlideNotes(slides[0]!, 'one');
    setSlideNotes(slides[2]!, 'three');

    const entries = getAllNotes(pres);
    expect(entries.length).toBe(2);
    expect(entries[0]).toEqual({ slideIndex: 0, notes: 'one' });
    expect(entries[1]).toEqual({ slideIndex: 2, notes: 'three' });
  });

  it('drops slides whose notes are empty', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank')!;
    const slide = addSlide(pres, { layout });
    setSlideNotes(slide, 'temporary');
    expect(getAllNotes(pres).length).toBe(1);
    setSlideNotes(slide, '');
    expect(getAllNotes(pres)).toEqual([]);
  });
});
