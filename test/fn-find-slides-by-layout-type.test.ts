// findSlidesByLayoutType — slides sharing an OOXML layout type.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  findSlidesByLayoutType,
  getSlideLayout,
  getSlideLayoutType,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: findSlidesByLayoutType', () => {
  it('returns slides whose layout @type matches', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    addSlide(pres, { layout: blank });
    addSlide(pres, { layout: blank });

    const matches = findSlidesByLayoutType(pres, 'blank');
    expect(matches.length).toBeGreaterThanOrEqual(2);
    for (const slide of matches) {
      const layout = getSlideLayout(slide)!;
      expect(getSlideLayoutType(layout)).toBe('blank');
    }
  });

  it('returns empty for an unused type', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    expect(findSlidesByLayoutType(pres, '__no-such-type__')).toEqual([]);
  });
});
