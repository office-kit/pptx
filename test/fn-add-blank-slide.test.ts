// addBlankSlide — auto-pick the blank layout if available.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  getSlideCount,
  getSlideLayout,
  getSlideLayoutType,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: addBlankSlide', () => {
  it('appends a slide and picks the blank-typed layout when present', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const before = getSlideCount(pres);
    const slide = addBlankSlide(pres);
    expect(getSlideCount(pres)).toBe(before + 1);
    const layout = getSlideLayout(slide)!;
    expect(getSlideLayoutType(layout)).toBe('blank');
  });
});
