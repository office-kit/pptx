// findSlidesByLayoutName — slides sharing a layout.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  findSlidesByLayoutName,
  getSlideIndex,
  getSlideLayout,
  getSlideLayoutName,
  getSlides,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findSlidesByLayoutName', () => {
  it('returns slides whose layout name matches exactly', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    addSlide(pres, { layout: blank });

    const matches = findSlidesByLayoutName(pres, 'Blank');
    expect(matches.length).toBeGreaterThanOrEqual(2);
    for (const slide of matches) {
      const layout = getSlideLayout(slide)!;
      expect(getSlideLayoutName(layout)).toBe('Blank');
    }
    // Indices align with getSlides
    const indices = matches.map((s) => getSlideIndex(pres, s));
    for (const i of indices) expect(i).toBeGreaterThanOrEqual(0);
  });

  it('returns empty for an unknown name', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    expect(findSlidesByLayoutName(pres, 'no-such-layout')).toEqual([]);
  });
});
