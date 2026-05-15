// Speaker-notes authoring (notesSlide part).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getPackagePartNames,
  getSlideNotes,
  getSlides,
  loadPresentation,
  savePresentation,
  setSlideNotes,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('L3: getSlideNotes / setSlideNotes', () => {
  it('returns null when the slide has no notesSlide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getSlideNotes(getSlides(pres)[0]!)).toBeNull();
  });

  it('creates a notesSlide part on first setSlideNotes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideNotes(slide, 'Speaker note line 1\nLine 2');

    const partNames = getPackagePartNames(pres);
    expect(partNames).toContain('/ppt/notesSlides/notesSlide1.xml');
    expect(getSlideNotes(getSlides(pres)[0]!)).toBe('Speaker note line 1\nLine 2');
  });

  it('updates an existing notesSlide on subsequent setSlideNotes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideNotes(slide, 'first');
    setSlideNotes(getSlides(pres)[0]!, 'second');
    expect(getSlideNotes(getSlides(pres)[0]!)).toBe('second');

    const noteParts = getPackagePartNames(pres).filter((n) =>
      n.startsWith('/ppt/notesSlides/notesSlide'),
    );
    expect(noteParts.length).toBe(1);
  });

  it('round-trips through save / load', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    setSlideNotes(slides[0]!, 'Round-tripped notes');
    setSlideNotes(slides[1]!, 'Second-slide notes');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reloadedSlides = getSlides(reloaded);
    expect(getSlideNotes(reloadedSlides[0]!)).toBe('Round-tripped notes');
    expect(getSlideNotes(reloadedSlides[1]!)).toBe('Second-slide notes');
  });
});
