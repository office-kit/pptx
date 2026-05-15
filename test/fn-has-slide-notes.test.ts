// hasSlideNotes — predicate for "does this slide have non-empty notes?"

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  getSlides,
  hasSlideNotes,
  loadPresentation,
  setSlideNotes,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: hasSlideNotes', () => {
  it('flips true after setSlideNotes', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    const slide = getSlides(pres).at(-1)!;
    expect(hasSlideNotes(slide)).toBe(false);
    setSlideNotes(slide, 'hello');
    expect(hasSlideNotes(slide)).toBe(true);
  });

  it('returns false for a fresh slide without notes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    for (const slide of getSlides(pres)) {
      expect(hasSlideNotes(slide)).toBe(false);
    }
  });
});
