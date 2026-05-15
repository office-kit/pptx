// getSlidesByLayout — find slides bound to a specific layout.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  getSlideLayout,
  getSlideLayoutName,
  getSlides,
  getSlidesByLayout,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlidesByLayout', () => {
  it('returns every slide bound to a given layout', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const titleSlideLayout = findSlideLayout(pres, 'Title Slide')!;
    const titleContentLayout = findSlideLayout(pres, 'Title and Content')!;
    addSlide(pres, { layout: titleSlideLayout });
    addSlide(pres, { layout: titleContentLayout });
    addSlide(pres, { layout: titleSlideLayout });

    expect(getSlidesByLayout(pres, titleSlideLayout)).toHaveLength(2);
    expect(getSlidesByLayout(pres, titleContentLayout)).toHaveLength(1);

    // Round-trip — the result must still match after save → reload.
    const reloadedFn = await loadPresentation(await savePresentation(pres));
    const reloadedLayout = findSlideLayout(reloadedFn, 'Title Slide')!;
    expect(getSlidesByLayout(reloadedFn, reloadedLayout)).toHaveLength(2);

    // Cross-check via layout name.
    const reloaded2 = await loadPresentation(await savePresentation(pres));
    const titled = getSlides(reloaded2).filter((s) => {
      const l = getSlideLayout(s);
      return l && getSlideLayoutName(l) === 'Title Slide';
    });
    expect(titled.length).toBe(2);
  });

  it('returns empty when no slide uses the layout', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const blank = findSlideLayout(pres, 'Blank')!;
    expect(getSlidesByLayout(pres, blank)).toEqual([]);
  });
});
