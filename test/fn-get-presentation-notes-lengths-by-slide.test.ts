// `getPresentationNotesLengthsBySlide(pres)` — dense per-slide notes
// length array.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getPresentationNotesLengthsBySlide,
  getSlides,
  loadPresentation,
  setSlideNotes,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getPresentationNotesLengthsBySlide', () => {
  it('returns 0 for slides without notes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getPresentationNotesLengthsBySlide(pres)).toEqual([0, 0]);
  });

  it('reflects authored notes per slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [slideA, slideB] = getSlides(pres);
    setSlideNotes(slideA!, 'hello world');
    setSlideNotes(slideB!, 'hi');
    const arr = getPresentationNotesLengthsBySlide(pres);
    expect(arr[0]).toBe('hello world'.length);
    expect(arr[1]).toBe('hi'.length);
  });
});
