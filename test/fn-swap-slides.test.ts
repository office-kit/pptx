// swapSlides — exchange two slides by index.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  getSlideTitle,
  getSlides,
  loadPresentation,
  setSlideTitle,
  swapSlides,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const buildThree = async () => {
  const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
  const layout = findSlideLayout(pres, 'Title and Content')!;
  setSlideTitle(addSlide(pres, { layout }), 'A');
  setSlideTitle(addSlide(pres, { layout }), 'B');
  setSlideTitle(addSlide(pres, { layout }), 'C');
  return pres;
};

const titles = (pres: Awaited<ReturnType<typeof buildThree>>): string[] =>
  getSlides(pres).map((s) => getSlideTitle(s) ?? '');

describe('fn API: swapSlides', () => {
  it('swaps adjacent slides', async () => {
    const pres = await buildThree();
    expect(titles(pres)).toEqual(['A', 'B', 'C']);
    swapSlides(pres, 0, 1);
    expect(titles(pres)).toEqual(['B', 'A', 'C']);
  });

  it('swaps non-adjacent slides', async () => {
    const pres = await buildThree();
    swapSlides(pres, 0, 2);
    expect(titles(pres)).toEqual(['C', 'B', 'A']);
  });

  it('is a no-op when the indices are equal', async () => {
    const pres = await buildThree();
    const before = titles(pres);
    swapSlides(pres, 1, 1);
    expect(titles(pres)).toEqual(before);
  });

  it('throws on out-of-range indices', async () => {
    const pres = await buildThree();
    expect(() => swapSlides(pres, 0, 99)).toThrow(/out of range/);
    expect(() => swapSlides(pres, -1, 0)).toThrow(/out of range/);
  });
});
