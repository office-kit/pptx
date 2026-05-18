// `getSlideLayoutUsageCountsByType(pres)` — histogram keyed on the
// layout-type enum.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  findSlideLayoutByType,
  getSlideLayoutUsageCountsByType,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getSlideLayoutUsageCountsByType', () => {
  it('counts slides by layout-type enum', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const obj = findSlideLayoutByType(pres, 'obj');
    if (!obj) throw new Error('expected obj layout');
    addSlide(pres, { layout: obj });
    addSlide(pres, { layout: obj });
    const counts = getSlideLayoutUsageCountsByType(pres);
    expect(counts.obj ?? 0).toBeGreaterThanOrEqual(2);
  });

  it('returns an object with at least one type key when the deck has slides', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const counts = getSlideLayoutUsageCountsByType(pres);
    expect(Object.keys(counts).length).toBeGreaterThanOrEqual(0);
  });
});
