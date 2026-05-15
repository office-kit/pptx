// getSlidesWithNotes — slides carrying non-empty speaker notes.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  getSlideIndex,
  getSlidesWithNotes,
  loadPresentation,
  setSlideNotes,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlidesWithNotes', () => {
  it('returns an empty list when nothing has notes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getSlidesWithNotes(pres)).toEqual([]);
  });

  it('returns only slides with notes attached', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank')!;
    // Three slides; only the first and third get notes.
    setSlideNotes(addSlide(pres, { layout }), 'speaker prompt 1');
    addSlide(pres, { layout });
    setSlideNotes(addSlide(pres, { layout }), 'speaker prompt 2');

    const hits = getSlidesWithNotes(pres);
    expect(hits.length).toBe(2);
    const indices = hits.map((s) => getSlideIndex(pres, s)).sort();
    expect(indices).toEqual([0, 2]);
  });

  it('drops slides whose notes were cleared back to empty', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank')!;
    const a = addSlide(pres, { layout });
    setSlideNotes(a, 'temporary');
    expect(getSlidesWithNotes(pres).length).toBe(1);
    setSlideNotes(a, '');
    expect(getSlidesWithNotes(pres)).toEqual([]);
  });
});
