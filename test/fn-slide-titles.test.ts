// getSlideTitles — every slide's title in order, null for missing.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  getSlideTitles,
  loadPresentation,
  setSlideTitle,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideTitles', () => {
  it('returns the titles in document order', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content')!;
    setSlideTitle(addSlide(pres, { layout }), 'A');
    setSlideTitle(addSlide(pres, { layout }), 'B');
    setSlideTitle(addSlide(pres, { layout }), 'C');
    expect(getSlideTitles(pres)).toEqual(['A', 'B', 'C']);
  });

  it('reports null for slides without a title placeholder', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    addSlide(pres, { layout: blank });
    expect(getSlideTitles(pres)).toEqual([null, null]);
  });

  it('reports an empty array for a deck with zero slides', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    expect(getSlideTitles(pres)).toEqual([]);
  });
});
