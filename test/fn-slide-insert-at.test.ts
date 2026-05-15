// addSlideAt / duplicateSlideAt — indexed insertion.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideAt,
  duplicateSlideAt,
  findSlideLayout,
  getSlideIndex,
  getSlides,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: addSlideAt', () => {
  it('inserts the new slide at the requested index', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const slide = addSlideAt(pres, 0, { layout });
    expect(getSlideIndex(pres, slide)).toBe(0);
    expect(getSlides(pres).length).toBe(3);
  });

  it('clamps to the end when the index is out of range', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    const slide = addSlideAt(pres, 99, { layout });
    expect(getSlideIndex(pres, slide)).toBe(2);
  });
});

describe('fn API: duplicateSlideAt', () => {
  it('duplicates and places the duplicate at the requested index', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const source = getSlides(pres)[1]!;
    const dup = duplicateSlideAt(pres, 0, source);
    expect(getSlideIndex(pres, dup)).toBe(0);
    expect(getSlides(pres).length).toBe(3);
  });
});
