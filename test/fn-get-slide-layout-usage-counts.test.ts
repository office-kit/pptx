// `getSlideLayoutUsageCounts(pres)` — name → number-of-slides
// histogram, including unreferenced layouts (count = 0).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayout,
  getSlideLayoutUsageCounts,
  getSlideLayouts,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideLayoutUsageCounts', () => {
  it('every layout in the package appears as a key', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const counts = getSlideLayoutUsageCounts(pres);
    const expected = getSlideLayouts(pres).length;
    expect(Object.keys(counts).length).toBe(expected);
  });

  it('counts every slide reference', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Title and Content');
    if (!layout) throw new Error('expected Title and Content layout');
    addSlide(pres, { layout });
    addSlide(pres, { layout });
    const counts = getSlideLayoutUsageCounts(pres);
    expect(counts['Title and Content']).toBe(2);
  });

  it('unreferenced layouts come back with count 0', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const counts = getSlideLayoutUsageCounts(pres);
    // The blank fixture has many layouts but only one slide that
    // references one specific layout; at least one layout must have a
    // count of 0.
    const zeros = Object.entries(counts).filter(([, n]) => n === 0);
    expect(zeros.length).toBeGreaterThan(0);
  });
});
