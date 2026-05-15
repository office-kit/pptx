// replaceTextInSlideNotes — slide-scoped notes replacer.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  getSlideNotes,
  getSlides,
  loadPresentation,
  replaceTextInSlideNotes,
  setSlideNotes,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: replaceTextInSlideNotes', () => {
  it('returns true and rewrites notes on a match', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    const slide = getSlides(pres).at(-1)!;
    setSlideNotes(slide, 'hello world');
    expect(replaceTextInSlideNotes(slide, 'world', 'there')).toBe(true);
    expect(getSlideNotes(getSlides(pres).at(-1)!)).toBe('hello there');
  });

  it('returns false when no notes or no match', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    const slide = getSlides(pres).at(-1)!;
    // No notes yet
    expect(replaceTextInSlideNotes(slide, 'x', 'y')).toBe(false);
    setSlideNotes(slide, 'plain');
    expect(replaceTextInSlideNotes(getSlides(pres).at(-1)!, 'no-such', 'y')).toBe(false);
  });
});
