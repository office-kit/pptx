// findSlidesByNotes — match slides by speaker-notes content.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  findSlidesByNotes,
  getSlideIndex,
  getSlides,
  loadPresentation,
  setSlideNotes,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findSlidesByNotes', () => {
  it('matches by substring', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    addSlide(pres, { layout: blank });
    const slides = getSlides(pres);
    setSlideNotes(slides[0]!, 'follow up with reviewer about TODO item');
    setSlideNotes(slides[1]!, 'done — ship it');
    const todoIdx = 0;
    const matches = findSlidesByNotes(pres, 'TODO');
    const indices = matches.map((s) => getSlideIndex(pres, s));
    expect(indices).toContain(todoIdx);
    expect(matches.length).toBe(1);
  });

  it('matches by RegExp and skips notes-less slides', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    addSlide(pres, { layout: blank });
    const slides = getSlides(pres);
    setSlideNotes(slides[0]!, 'meeting 2026-05-15');
    // slide 1 has no notes
    expect(findSlidesByNotes(pres, /\d{4}-\d{2}-\d{2}/).length).toBe(1);
    expect(findSlidesByNotes(pres, /^x/).length).toBe(0);
  });
});
